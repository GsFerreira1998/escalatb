from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from datetime import datetime
import pandas as pd
import os
from sistema_escala import SistemaEscalaExcel
import io

app = Flask(__name__)
app.secret_key = 'escala_rodizio_secreto_2024'
sistema = SistemaEscalaExcel()

@app.route('/')
def index():
    """Página inicial"""
    # Carregar escalas existentes
    arquivos = sistema.listar_escalas_existentes(return_list=True)
    
    # Estatísticas rápidas
    stats = {
        'total_escalas': len(arquivos),
        'ultimas_escalas': arquivos[:5] if arquivos else [],
        'total_funcionarios': sum(len(v) for v in sistema.funcionarios.values())
    }
    
    return render_template('index.html', stats=stats)

@app.route('/gerar_escala', methods=['GET', 'POST'])
def gerar_escala():
    """Gerar nova escala mensal"""
    if request.method == 'POST':
        try:
            ano = int(request.form['ano'])
            mes = int(request.form['mes'])
            semanas = int(request.form.get('semanas', 4))
            
            # Verificar se já existe
            arquivo_existente = f"{sistema.diretorio_escalas}/ESCALA_{ano}_{mes:02d}.xlsx"
            existe = os.path.exists(arquivo_existente)
            
            if existe and 'confirmar' not in request.form:
                # Mostrar confirmação
                return render_template('gerar_escala.html', 
                                     confirmar=True, 
                                     ano=ano, 
                                     mes=mes, 
                                     semanas=semanas,
                                     existe=True)
            
            # Gerar escala
            df_escala = sistema.gerar_escala_mensal(ano, mes, semanas)
            
            # Salvar
            arquivo_salvo = sistema.salvar_escala_excel(df_escala, ano, mes)
            
            # Verificar regras
            verificacao = sistema.verificar_regras(df_escala)
            
            # Calcular estatísticas
            stats = {
                'arquivo': arquivo_salvo,
                'total_funcionarios': len(df_escala['Funcionário'].unique()),
                'semanas': semanas,
                'erros': len(verificacao['erros']),
                'regras_ok': sum([
                    verificacao['regra_5_dias'],
                    verificacao['regra_folgas_seguidas'],
                    verificacao['regra_fim_semana_seguido'],
                    verificacao['cobertura_sabado'],
                    verificacao['cobertura_domingo'],
                    verificacao['rodizio_domingo'],
                    verificacao['rodizio_sabado'],
                    verificacao['rodizio_folgas']
                ]),
                'primeiros_funcionarios': df_escala.head(10).to_dict('records')
            }
            
            flash(f'Escala gerada com sucesso para {mes:02d}/{ano}!', 'success')
            return render_template('gerar_escala_resultado.html', 
                                 stats=stats, 
                                 erros=verificacao['erros'][:10],
                                 ano=ano, 
                                 mes=mes)
            
        except Exception as e:
            flash(f'Erro ao gerar escala: {str(e)}', 'danger')
            return redirect('/gerar_escala')
    
    return render_template('gerar_escala.html', confirmar=False)

@app.route('/listar_escalas')
def listar_escalas():
    """Listar todas as escalas existentes"""
    arquivos = sistema.listar_escalas_existentes(return_list=True)
    
    escalas_detalhadas = []
    for arquivo in arquivos:
        partes = arquivo.replace('ESCALA_', '').replace('.xlsx', '').split('_')
        if len(partes) >= 2:
            ano, mes = partes[0], partes[1]
            caminho = f"{sistema.diretorio_escalas}/{arquivo}"
            tamanho = f"{os.path.getsize(caminho) / 1024:.1f} KB" if os.path.exists(caminho) else "N/A"
            
            escalas_detalhadas.append({
                'arquivo': arquivo,
                'ano': ano,
                'mes': mes,
                'tamanho': tamanho,
                'caminho': caminho
            })
    
    return render_template('listar_escalas.html', escalas=escalas_detalhadas)

@app.route('/visualizar_escala/<ano>/<mes>')
def visualizar_escala(ano, mes):
    """Visualizar escala específica"""
    try:
        arquivo = f"{sistema.diretorio_escalas}/ESCALA_{ano}_{int(mes):02d}.xlsx"
        
        if not os.path.exists(arquivo):
            flash('Escala não encontrada', 'danger')
            return redirect('/listar_escalas')
        
        # Carregar dados
        df_escala = pd.read_excel(arquivo, sheet_name='ESCALA_COMPLETA')
        
        # Converter para lista de dicionários para o template
        dados = df_escala.head(50).to_dict('records')
        
        # Estatísticas
        stats = {
            'ano': ano,
            'mes': mes,
            'total_registros': len(df_escala),
            'semanas': df_escala['Semana do Mês'].nunique(),
            'funcionarios': df_escala['Funcionário'].nunique(),
            'total_dias_trabalhados': int(df_escala['Dias Trabalhados'].sum()),
            'dados': dados
        }
        
        return render_template('visualizar.html', stats=stats)
        
    except Exception as e:
        flash(f'Erro ao carregar escala: {str(e)}', 'danger')
        return redirect('/listar_escalas')

@app.route('/download_escala/<ano>/<mes>')
def download_escala(ano, mes):
    """Download do arquivo Excel"""
    try:
        arquivo = f"{sistema.diretorio_escalas}/ESCALA_{ano}_{int(mes):02d}.xlsx"
        
        if not os.path.exists(arquivo):
            flash('Arquivo não encontrado', 'danger')
            return redirect('/listar_escalas')
        
        return send_file(
            arquivo,
            as_attachment=True,
            download_name=f"ESCALA_{ano}_{int(mes):02d}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        flash(f'Erro ao baixar arquivo: {str(e)}', 'danger')
        return redirect('/listar_escalas')

@app.route('/contadores', methods=['GET', 'POST'])
def contadores():
    """Ver contadores de fim de semana"""
    if request.method == 'POST':
        try:
            ano = int(request.form['ano'])
            mes = int(request.form['mes'])
            
            if mes == 0:  # Ano inteiro
                arquivos_ano = []
                for m in range(1, 13):
                    arquivo = f"{sistema.diretorio_escalas}/ESCALA_{ano}_{m:02d}.xlsx"
                    if os.path.exists(arquivo):
                        arquivos_ano.append(arquivo)
                
                if not arquivos_ano:
                    flash(f'Nenhuma escala encontrada para o ano {ano}', 'warning')
                    return render_template('contadores.html')
                
                # Combinar contadores
                todos_contadores = []
                for arquivo in arquivos_ano:
                    try:
                        df = pd.read_excel(arquivo, sheet_name='CONTADORES_FIM_SEMANA')
                        todos_contadores.append(df)
                    except:
                        continue
                
                if not todos_contadores:
                    flash('Não foi possível carregar contadores', 'warning')
                    return render_template('contadores.html')
                
                df_todos = pd.concat(todos_contadores, ignore_index=True)
                df_agrupado = df_todos.groupby('Funcionário').agg({
                    'Sábados Trabalhados': 'sum',
                    'Domingos Trabalhados': 'sum',
                    'Total Fim de Semana': 'sum',
                    'Rodada Domingo': 'max',
                    'Rodada Sábado': 'max'
                }).reset_index()
                
                contadores = df_agrupado.sort_values('Domingos Trabalhados').head(50).to_dict('records')
                titulo = f"Contadores Acumulados do Ano {ano}"
                
            else:  # Mês específico
                arquivo = f"{sistema.diretorio_escalas}/ESCALA_{ano}_{mes:02d}.xlsx"
                
                if not os.path.exists(arquivo):
                    flash(f'Escala não encontrada para {mes:02d}/{ano}', 'warning')
                    return render_template('contadores.html')
                
                df_contadores = pd.read_excel(arquivo, sheet_name='CONTADORES_FIM_SEMANA')
                contadores = df_contadores.sort_values('Domingos Trabalhados').head(50).to_dict('records')
                titulo = f"Contadores de {mes:02d}/{ano}"
            
            return render_template('contadores.html', 
                                 contadores=contadores, 
                                 titulo=titulo,
                                 ano=ano, 
                                 mes=mes)
            
        except Exception as e:
            flash(f'Erro ao carregar contadores: {str(e)}', 'danger')
    
    return render_template('contadores.html')

@app.route('/gerar_relatorio', methods=['GET', 'POST'])
def gerar_relatorio():
    """Gerar relatório anual"""
    if request.method == 'POST':
        try:
            ano = int(request.form['ano'])
            arquivo_relatorio = sistema.gerar_relatorio_anual(int(ano))
            
            if arquivo_relatorio:
                flash(f'Relatório anual {ano} gerado com sucesso!', 'success')
                return render_template('relatorio.html', 
                                     sucesso=True, 
                                     arquivo=arquivo_relatorio,
                                     ano=ano)
            else:
                flash(f'Não foi possível gerar relatório para {ano}', 'warning')
                
        except Exception as e:
            flash(f'Erro ao gerar relatório: {str(e)}', 'danger')
    
    return render_template('relatorio.html')

@app.route('/verificar_rodizio', methods=['GET', 'POST'])
def verificar_rodizio():
    """Verificar rodízio perfeito"""
    if request.method == 'POST':
        try:
            ano = int(request.form['ano'])
            mes = int(request.form['mes'])
            
            arquivo = f"{sistema.diretorio_escalas}/ESCALA_{ano}_{int(mes):02d}.xlsx"
            
            if not os.path.exists(arquivo):
                flash(f'Escala não encontrada para {mes:02d}/{ano}', 'warning')
                return render_template('rodizio.html')
            
            df_escala = pd.read_excel(arquivo, sheet_name='ESCALA_COMPLETA')
            rodizio = sistema.verificar_rodizio_perfeito(df_escala)
            
            # Carregar contadores para detalhes
            df_contadores = pd.read_excel(arquivo, sheet_name='CONTADORES_FIM_SEMANA')
            
            # Estatísticas por ilha
            stats_ilha = []
            for ilha in sistema.funcionarios.keys():
                if ilha in rodizio['rodizio_domingo_por_ilha']:
                    domingos = rodizio['rodizio_domingo_por_ilha'][ilha]
                    sabados = rodizio['rodizio_sabado_por_ilha'][ilha]
                    
                    min_dom = min(domingos.values())
                    max_dom = max(domingos.values())
                    min_sab = min(sabados.values())
                    max_sab = max(sabados.values())
                    
                    stats_ilha.append({
                        'ilha': ilha,
                        'dom_min': min_dom,
                        'dom_max': max_dom,
                        'dom_diff': max_dom - min_dom,
                        'sab_min': min_sab,
                        'sab_max': max_sab,
                        'sab_diff': max_sab - min_sab
                    })
            
            top_menos_domingos = df_contadores.nsmallest(5, 'Domingos Trabalhados').to_dict('records')
            
            return render_template('rodizio.html',
                                 rodizio=rodizio,
                                 stats_ilha=stats_ilha,
                                 top_menos_domingos=top_menos_domingos,
                                 ano=ano,
                                 mes=mes)
            
        except Exception as e:
            flash(f'Erro ao verificar rodízio: {str(e)}', 'danger')
    
    return render_template('rodizio.html')

@app.route('/verificar_disponibilidade', methods=['GET', 'POST'])
def verificar_disponibilidade():
    """Verificar disponibilidade por data"""
    if request.method == 'POST':
        try:
            ano = int(request.form['ano'])
            mes = int(request.form['mes'])
            semana = int(request.form['semana'])
            dia = request.form['dia'].lower()
            
            # Mapear dia para coluna
            mapa_dias = {
                'seg': 'Seg', 'ter': 'Ter', 'qua': 'Qua', 'qui': 'Qui',
                'sex': 'Sex', 'sab': 'Sáb', 'dom': 'Dom'
            }
            
            if dia not in mapa_dias:
                flash('Dia inválido', 'danger')
                return render_template('disponibilidade.html')
            
            coluna_dia = mapa_dias[dia]
            arquivo = f"{sistema.diretorio_escalas}/ESCALA_{ano}_{mes:02d}.xlsx"
            
            if not os.path.exists(arquivo):
                flash(f'Escala não encontrada para {mes:02d}/{ano}', 'warning')
                return render_template('disponibilidade.html')
            
            df_escala = pd.read_excel(arquivo, sheet_name='ESCALA_COMPLETA')
            df_dia = df_escala[df_escala['Semana do Mês'] == semana].copy()
            
            disponiveis = df_dia[df_dia[coluna_dia] == 'P']
            folga = df_dia[df_dia[coluna_dia] == 'F']
            
            # Agrupar por ilha
            disponiveis_por_ilha = []
            for ilha in disponiveis['Ilha'].unique():
                funcs = disponiveis[disponiveis['Ilha'] == ilha]['Funcionário'].tolist()
                disponiveis_por_ilha.append({
                    'ilha': ilha,
                    'quantidade': len(funcs),
                    'funcionarios': [{'nome': f.split()[0] + ' ' + f.split()[1], 'completo': f} for f in funcs[:5]]
                })
            
            amostra_disponiveis = disponiveis.head(10).to_dict('records')
            
            return render_template('disponibilidade.html',
                                 disponiveis_por_ilha=disponiveis_por_ilha,
                                 amostra_disponiveis=amostra_disponiveis,
                                 total=len(df_dia),
                                 disponiveis=len(disponiveis),
                                 folga=len(folga),
                                 ano=ano,
                                 mes=mes,
                                 semana=semana,
                                 dia=dia)
            
        except Exception as e:
            flash(f'Erro ao verificar disponibilidade: {str(e)}', 'danger')
    
    return render_template('disponibilidade.html')

@app.route('/rodizio_folgas', methods=['GET', 'POST'])
def rodizio_folgas():
    """Verificar rodízio de folgas"""
    if request.method == 'POST':
        try:
            ano = int(request.form['ano'])
            mes = int(request.form['mes'])
            
            arquivo = f"{sistema.diretorio_escalas}/ESCALA_{ano}_{int(mes):02d}.xlsx"
            
            if not os.path.exists(arquivo):
                flash(f'Escala não encontrada para {mes:02d}/{ano}', 'warning')
                return render_template('rodizio_folgas.html')
            
            # Carregar dados de folgas
            df_folgas = pd.read_excel(arquivo, sheet_name='RODÍZIO_FOLGAS')
            df_estatisticas = pd.read_excel(arquivo, sheet_name='ESTAT_FOLGAS')
            
            # Encontrar funcionários desbalanceados
            desbalanceados = []
            for _, row in df_folgas.iterrows():
                folgas = [row['Folgas Segunda'], row['Folgas Terça'], row['Folgas Quarta'], 
                         row['Folgas Quinta'], row['Folgas Sexta']]
                max_folgas = max(folgas)
                min_folgas = min(folgas)
                
                if max_folgas - min_folgas > 2:
                    desbalanceados.append({
                        'funcionario': row['Funcionário'],
                        'ilha': row['Ilha'],
                        'folgas': folgas,
                        'diferenca': max_folgas - min_folgas
                    })
            
            estatisticas = df_estatisticas.to_dict('records')
            
            return render_template('rodizio_folgas.html',
                                 estatisticas=estatisticas,
                                 desbalanceados=desbalanceados,
                                 total_desbalanceados=len(desbalanceados),
                                 ano=ano,
                                 mes=mes)
            
        except Exception as e:
            flash(f'Erro ao verificar folgas: {str(e)}', 'danger')
    
    return render_template('rodizio_folgas.html')

if __name__ == '__main__':
    # Criar diretórios necessários
    os.makedirs('templates', exist_ok=True)
    
    # Verificar se o sistema de escalas existe
    if not os.path.exists(sistema.diretorio_escalas):
        print(f"Criando diretório {sistema.diretorio_escalas}")
        os.makedirs(sistema.diretorio_escalas, exist_ok=True)
    
    app.run(debug=True, port=5000)