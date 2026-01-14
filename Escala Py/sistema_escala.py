import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import os
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import warnings
from collections import defaultdict, deque, Counter
warnings.filterwarnings('ignore')

class SistemaEscalaExcel:
    """
    Sistema completo de escala 5x2 usando apenas Excel para histórico
    COM RODÍZIO PERFEITO DE FIM DE SEMANA E RODÍZIO DE FOLGAS
    """
    
    def __init__(self):
        # Dados dos funcionários por ilha
        self.funcionarios = {
            "ILHA SP": [
                "ERNANE ALVES BRITO", "JULIA SANTOS SIQUEIRA", "NICOLI PEREIRA DA SILVA",
                "MARCO ANTONIO DE CAMPOS", "JOBSON DA SILVA BRITO", "CARLA ROBERTA ALVES",
                "LUCAS SANTOS", "LEONEL PEREIRA", "DANIELA ALVES DE CARVALHO"
            ],
            "ILHA SC": [
                "ROBSON DE LIMA SANTOS", "AMANDA ALVES DE SOUZA", "LETICIA SILVA DE SANTANA",
                "ALAN DE SOUZA CHAVES", "VICTOR HUGO DE SOUZA VASCONSELLOS",
                "MATEUS FAGUNDES DE LIMA", "MANUELA MENESES MACHADO"
            ],
            "ILHA NO": [
                "UENDERSON ENIS GOMES PEREIRA", "VINICIUS HENRIQUE VIANA DOS SANTOS",
                "MARCELO ANTONIO DA CRUZ", "LEANDRO BATISTA DE MELLO",
                "MATEUS DA SILVA CAMPOS", "DAPHNE FELIPPE DA HORA",
                "FELIPE ALEXANDRE DE ALMEIDA", "MARCOS ANDRE OLIVA TEXEIRA"
            ],
            "ILHA SU": [
                "RODRIGO PEREIRA MARQUES MACEGOZA", "FERNANDA ALVES DE BRITO",
                "TALITA MARTINS PAZ", "EDUARDO GARCIA MASSA", "MATHEUS DOS SANTOS SILVA",
                "FRANCISCO LUTHELLE CARNEIRO SEVERO", "AUGUSTO PERNHA SANTOS",
                "RAY BORAZO", "THIAGO CASTRO"
            ]
        }
        
        self.dias_semana = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
        self.dias_completos = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
        
        # Diretório para salvar as escalas
        self.diretorio_escalas = "ESCALAS_HISTORICO"
        os.makedirs(self.diretorio_escalas, exist_ok=True)
        
        # Sistema de rodízio por ilha
        self.rodizio_ilhas = {}
        
        # Sistema de rodízio de folgas (novo)
        self.rodizio_folgas = {}
        
        self.inicializar_rodizio()
    
    def inicializar_rodizio(self):
        """Inicializa o sistema de rodízio para cada ilha"""
        for ilha, lista_func in self.funcionarios.items():
            # Para cada ilha, temos duas filas:
            # 1. fila_domingo: Quem ainda NÃO pegou domingo (ou pegou menos que outros)
            # 2. fila_sabado: Quem ainda NÃO pegou sábado (ou pegou menos que outros)
            self.rodizio_ilhas[ilha] = {
                'fila_domingo': deque(lista_func),  # Começa com todos
                'fila_sabado': deque(lista_func),   # Começa com todos
                'domingos_pegos': {func: 0 for func in lista_func},
                'sabados_pegos': {func: 0 for func in lista_func},
                'rodada_domingo': 0,  # Contador de rodadas completas
                'rodada_sabado': 0,   # Contador de rodadas completas
                'ultimo_domingo': None,  # Quem pegou o último domingo
                'ultimo_sabado': None    # Quem pegou o último sábado
            }
            
            # Inicializar rodízio de folgas
            self.rodizio_folgas[ilha] = {
                'contador_folgas': {func: {i: 0 for i in range(5)} for func in lista_func},  # 0=Seg, 1=Ter, 2=Qua, 3=Qui, 4=Sex
                'ultimas_folgas': {func: [] for func in lista_func},  # Histórico das últimas folgas
                'sequencia_folgas': {func: deque() for func in lista_func},  # Sequência de folgas a serem distribuídas
                'prioridade_folgas': deque(lista_func)  # Quem deve pegar folgas primeiro
            }
    
    def carregar_contadores_mes_anterior(self, ano: int, mes: int) -> Dict:
        """
        Carrega os contadores de fim de semana do mês anterior a partir do Excel
        
        Args:
            ano: Ano atual
            mes: Mês atual
            
        Returns:
            Dicionário com contadores de cada funcionário
        """
        # Calcular mês anterior
        if mes == 1:
            ano_anterior = ano - 1
            mes_anterior = 12
        else:
            ano_anterior = ano
            mes_anterior = mes - 1
        
        # Nome do arquivo do mês anterior
        arquivo_anterior = f"{self.diretorio_escalas}/ESCALA_{ano_anterior}_{mes_anterior:02d}.xlsx"
        
        contadores = {}
        
        # Se existe arquivo do mês anterior, carrega os contadores
        if os.path.exists(arquivo_anterior):
            try:
                # Carregar a aba de contadores
                df_contadores = pd.read_excel(arquivo_anterior, sheet_name='CONTADORES_FIM_SEMANA')
                
                for _, row in df_contadores.iterrows():
                    funcionario = row['Funcionário']
                    contadores[funcionario] = {
                        'sabados_trabalhados': row['Sábados Trabalhados'],
                        'domingos_trabalhados': row['Domingos Trabalhados'],
                        'total_fim_semana': row['Total Fim de Semana'],
                        'rodada_domingo': row.get('Rodada Domingo', 0),
                        'rodada_sabado': row.get('Rodada Sábado', 0)
                    }
                
                print(f"✅ Contadores carregados de {mes_anterior:02d}/{ano_anterior}")
                
                # Reconstruir o sistema de rodízio com os dados históricos
                self.reconstruir_rodizio(contadores)
                
                return contadores
                
            except Exception as e:
                print(f"⚠️  Não foi possível carregar contadores do mês anterior: {e}")
        
        # Se não encontrou arquivo anterior, inicia contadores zerados
        print("📭 Iniciando com contadores zerados (primeiro mês ou mês anterior não encontrado)")
        
        # Inicializar contadores zerados para todos os funcionários
        for ilha, lista_func in self.funcionarios.items():
            for funcionario in lista_func:
                contadores[funcionario] = {
                    'sabados_trabalhados': 0,
                    'domingos_trabalhados': 0,
                    'total_fim_semana': 0,
                    'rodada_domingo': 0,
                    'rodada_sabado': 0
                }
        
        return contadores
    
    def reconstruir_rodizio(self, contadores: Dict):
        """
        Reconstrói o sistema de rodízio com base nos contadores históricos
        
        Args:
            contadores: Dicionário com contadores históricos
        """
        # Para cada ilha
        for ilha, lista_func in self.funcionarios.items():
            # Ordenar funcionários por quem tem MENOS domingos
            funcs_ordenados_dom = sorted(
                lista_func,
                key=lambda f: contadores.get(f, {}).get('domingos_trabalhados', 0)
            )
            
            # Ordenar funcionários por quem tem MENOS sábados
            funcs_ordenados_sab = sorted(
                lista_func,
                key=lambda f: contadores.get(f, {}).get('sabados_trabalhados', 0)
            )
            
            # Atualizar as filas
            self.rodizio_ilhas[ilha]['fila_domingo'] = deque(funcs_ordenados_dom)
            self.rodizio_ilhas[ilha]['fila_sabado'] = deque(funcs_ordenados_sab)
            
            # Atualizar contadores de rodadas
            # Rodada = quantas vezes todos pegaram pelo menos 1
            min_domingos = min(contadores.get(f, {}).get('domingos_trabalhados', 0) for f in lista_func)
            min_sabados = min(contadores.get(f, {}).get('sabados_trabalhados', 0) for f in lista_func)
            
            self.rodizio_ilhas[ilha]['rodada_domingo'] = min_domingos
            self.rodizio_ilhas[ilha]['rodada_sabado'] = min_sabados
            
            # Atualizar contadores individuais
            for func in lista_func:
                self.rodizio_ilhas[ilha]['domingos_pegos'][func] = contadores.get(func, {}).get('domingos_trabalhados', 0)
                self.rodizio_ilhas[ilha]['sabados_pegos'][func] = contadores.get(func, {}).get('sabados_trabalhados', 0)
        
        print("✅ Sistema de rodízio reconstruído com base no histórico")
    
    def obter_proximo_domingo(self, ilha: str) -> str:
        """
        Obtém o próximo funcionário que deve trabalhar no domingo
        seguindo a regra: só pode pegar novamente se TODOS já pegaram
        
        Args:
            ilha: Nome da ilha
            
        Returns:
            Nome do funcionário
        """
        rodizio = self.rodizio_ilhas[ilha]
        
        # VERIFICAÇÃO FORTE: Verificar se TODOS já pegaram pelo menos 1 domingo
        todos_tem_domingo = all(v > 0 for v in rodizio['domingos_pegos'].values())
        
        if not todos_tem_domingo:
            # Se ainda não, pegar apenas quem tem 0 domingos
            candidatos = [f for f, v in rodizio['domingos_pegos'].items() if v == 0]
            
            if not candidatos:
                # Se não há quem tenha 0, pegar quem tem menos
                min_domingos = min(rodizio['domingos_pegos'].values())
                candidatos = [f for f, v in rodizio['domingos_pegos'].items() if v == min_domingos]
            
            # Ordenar pela fila
            for func in rodizio['fila_domingo']:
                if func in candidatos:
                    funcionario = func
                    break
        else:
            # Se todos já pegaram, pegar quem tem menos domingos
            min_domingos = min(rodizio['domingos_pegos'].values())
            candidatos = [f for f, v in rodizio['domingos_pegos'].items() if v == min_domingos]
            
            # Ordenar pela fila
            for func in rodizio['fila_domingo']:
                if func in candidatos:
                    funcionario = func
                    break
            else:
                funcionario = candidatos[0]
        
        # Atualizar fila (remover o escolhido e colocar no final)
        if funcionario in rodizio['fila_domingo']:
            rodizio['fila_domingo'].remove(funcionario)
        rodizio['fila_domingo'].append(funcionario)
        
        # Incrementar contador
        rodizio['domingos_pegos'][funcionario] += 1
        rodizio['ultimo_domingo'] = funcionario
        
        # Verificar se completou uma rodada
        if todos_tem_domingo and all(v == rodizio['domingos_pegos'][funcionario] for v in rodizio['domingos_pegos'].values()):
            rodizio['rodada_domingo'] += 1
            print(f"    🎯 {ilha}: COMPLETOU RODADA DE DOMINGO {rodizio['rodada_domingo']}")
        
        return funcionario
    
    def obter_proximo_sabado(self, ilha: str) -> str:
        """
        Obtém o próximo funcionário que deve trabalhar no sábado
        seguindo a regra: só pode pegar novamente se TODOS já pegaram
        
        Args:
            ilha: Nome da ilha
            
        Returns:
            Nome do funcionário
        """
        rodizio = self.rodizio_ilhas[ilha]
        
        # VERIFICAÇÃO FORTE: Verificar se TODOS já pegaram pelo menos 1 sábado
        todos_tem_sabado = all(v > 0 for v in rodizio['sabados_pegos'].values())
        
        if not todos_tem_sabado:
            # Se ainda não, pegar apenas quem tem 0 sábados
            candidatos = [f for f, v in rodizio['sabados_pegos'].items() if v == 0]
            
            if not candidatos:
                # Se não há quem tenha 0, pegar quem tem menos
                min_sabados = min(rodizio['sabados_pegos'].values())
                candidatos = [f for f, v in rodizio['sabados_pegos'].items() if v == min_sabados]
            
            # Ordenar pela fila
            for func in rodizio['fila_sabado']:
                if func in candidatos:
                    funcionario = func
                    break
        else:
            # Se todos já pegaram, pegar quem tem menos sábados
            min_sabados = min(rodizio['sabados_pegos'].values())
            candidatos = [f for f, v in rodizio['sabados_pegos'].items() if v == min_sabados]
            
            # Ordenar pela fila
            for func in rodizio['fila_sabado']:
                if func in candidatos:
                    funcionario = func
                    break
            else:
                funcionario = candidatos[0]
        
        # Atualizar fila (remover o escolhido e colocar no final)
        if funcionario in rodizio['fila_sabado']:
            rodizio['fila_sabado'].remove(funcionario)
        rodizio['fila_sabado'].append(funcionario)
        
        # Incrementar contador
        rodizio['sabados_pegos'][funcionario] += 1
        rodizio['ultimo_sabado'] = funcionario
        
        # Verificar se completou uma rodada
        if todos_tem_sabado and all(v == rodizio['sabados_pegos'][funcionario] for v in rodizio['sabados_pegos'].values()):
            rodizio['rodada_sabado'] += 1
            print(f"    🎯 {ilha}: COMPLETOU RODADA DE SÁBADO {rodizio['rodada_sabado']}")
        
        return funcionario
    
    def obter_melhor_folga_semanal(self, ilha: str, funcionario: str, dias_disponiveis: List[int]) -> int:
        """
        Determina o melhor dia para folga baseado no rodízio histórico
        
        Args:
            ilha: Nome da ilha
            funcionario: Nome do funcionário
            dias_disponiveis: Lista de índices de dias disponíveis para folga (0=Seg, 4=Sex)
            
        Returns:
            Índice do dia escolhido para folga
        """
        rodizio = self.rodizio_folgas[ilha]
        
        # Se não tem histórico, começar com segunda
        if not rodizio['ultimas_folgas'][funcionario]:
            # Escolher o primeiro dia disponível
            return dias_disponiveis[0]
        
        # Verificar qual dia da semana a pessoa tem MENOS folgas
        contadores = rodizio['contador_folgas'][funcionario]
        dias_ordenados = sorted(dias_disponiveis, key=lambda d: contadores[d])
        
        # Pegar o dia com menos folgas
        melhor_dia = dias_ordenados[0]
        
        # Verificar se a pessoa tem uma sequência de folgas
        if rodizio['sequencia_folgas'][funcionario]:
            proxima_folga = rodizio['sequencia_folgas'][funcionario][0]
            if proxima_folga in dias_disponiveis:
                melhor_dia = proxima_folga
                rodizio['sequencia_folgas'][funcionario].popleft()
        
        # Atualizar contador
        rodizio['contador_folgas'][funcionario][melhor_dia] += 1
        rodizio['ultimas_folgas'][funcionario].append(melhor_dia)
        
        # Manter apenas as últimas 10 folgas
        if len(rodizio['ultimas_folgas'][funcionario]) > 10:
            rodizio['ultimas_folgas'][funcionario].pop(0)
        
        return melhor_dia
    
    def gerar_escala_mensal(self, ano: int, mes: int, semanas: int = 4) -> pd.DataFrame:
        """
        Gera escala para um mês específico usando RODÍZIO PERFEITO
        
        Args:
            ano: Ano da escala
            mes: Mês da escala
            semanas: Número de semanas
            
        Returns:
            DataFrame com a escala mensal
        """
        print(f"\n📊 GERANDO ESCALA PARA {mes:02d}/{ano}")
        print("=" * 50)
        print("📋 REGRAS DO RODÍZIO:")
        print("   1. Domingo: Só pode pegar novamente se TODOS da ilha já pegaram")
        print("   2. Sábado: Só pode pegar novamente se TODOS da ilha já pegaram")
        print("   3. Folgas: Rodízio entre segunda e sexta para equalizar")
        print("   4. Prioridade: Domingo > Sábado > Folgas")
        
        # Carregar contadores do mês anterior
        contadores = self.carregar_contadores_mes_anterior(ano, mes)
        
        # Mostrar estado atual do rodízio
        print(f"\n📈 ESTADO DO RODÍZIO:")
        for ilha in self.rodizio_ilhas:
            rodizio = self.rodizio_ilhas[ilha]
            min_dom = min(rodizio['domingos_pegos'].values())
            max_dom = max(rodizio['domingos_pegos'].values())
            min_sab = min(rodizio['sabados_pegos'].values())
            max_sab = max(rodizio['sabados_pegos'].values())
            
            diff_dom = max_dom - min_dom
            diff_sab = max_sab - min_sab
            
            status_dom = "✅" if diff_dom <= 1 else "⚠️"
            status_sab = "✅" if diff_sab <= 1 else "⚠️"
            
            print(f"   {ilha}: Dom [{min_dom}-{max_dom}] {status_dom} | Sáb [{min_sab}-{max_sab}] {status_sab} | R:D{rodizio['rodada_domingo']} S:{rodizio['rodada_sabado']}")
        
        # Lista para armazenar todas as semanas
        todas_escalas = []
        
        # Contadores temporários para este mês
        contadores_mes_atual = {}
        for func in contadores:
            contadores_mes_atual[func] = {
                'sabados': 0,
                'domingos': 0,
                'total': 0
            }
        
        # Gerar cada semana do mês
        for semana_num in range(1, semanas + 1):
            print(f"\n  📅 Gerando semana {semana_num}/{semanas}...")
            
            escala_semanal = self.gerar_escala_semanal_rodizio(
                semana_num=semana_num,
                contadores=contadores,
                contadores_mes_atual=contadores_mes_atual,
                funcionarios=self.funcionarios
            )
            
            # Adicionar colunas de identificação
            escala_semanal['Ano'] = ano
            escala_semanal['Mês'] = mes
            escala_semanal['Semana do Mês'] = semana_num
            
            todas_escalas.append(escala_semanal)
            
            # Mostrar distribuição desta semana
            self.mostrar_distribuicao_semana(escala_semanal, semana_num)
        
        # Combinar todas as semanas
        df_escala_mensal = pd.concat(todas_escalas, ignore_index=True)
        
        # Reordenar colunas
        colunas = ['Ano', 'Mês', 'Semana do Mês', 'Funcionário', 'Ilha'] + \
                  self.dias_semana + ['Dias Trabalhados', 'Folgas']
        df_escala_mensal = df_escala_mensal[colunas]
        
        # Atualizar contadores com os dados do mês atual
        for _, row in df_escala_mensal.iterrows():
            func = row['Funcionário']
            if func in contadores:
                # Atualizar contadores básicos
                contadores[func]['sabados_trabalhados'] += contadores_mes_atual[func]['sabados']
                contadores[func]['domingos_trabalhados'] += contadores_mes_atual[func]['domingos']
                contadores[func]['total_fim_semana'] += contadores_mes_atual[func]['total']
                
                # Atualizar rodadas no contador
                for ilha, lista_func in self.funcionarios.items():
                    if func in lista_func:
                        contadores[func]['rodada_domingo'] = self.rodizio_ilhas[ilha]['rodada_domingo']
                        contadores[func]['rodada_sabado'] = self.rodizio_ilhas[ilha]['rodada_sabado']
                        break
        
        return df_escala_mensal
    
    def mostrar_distribuicao_semana(self, df_semana: pd.DataFrame, semana_num: int):
        """Mostra a distribuição de fins de semana para uma semana específica"""
        print(f"    📊 Distribuição semana {semana_num}:")
        
        for ilha in df_semana['Ilha'].unique():
            df_ilha = df_semana[df_semana['Ilha'] == ilha]
            
            # Contar quem trabalha no fim de semana
            sabado = df_ilha[df_ilha['Sáb'] == 'P']['Funcionário'].tolist()
            domingo = df_ilha[df_ilha['Dom'] == 'P']['Funcionário'].tolist()
            
            # Abreviar nomes
            sab_abreviados = [' '.join(f.split()[:2]) for f in sabado[:2]]
            dom_abreviados = [' '.join(f.split()[:2]) for f in domingo[:1]] if domingo else []
            
            print(f"      {ilha}: Sábado={sab_abreviados} | Domingo={dom_abreviados}")
    
    def gerar_escala_semanal_rodizio(self, semana_num: int, contadores: Dict, 
                                   contadores_mes_atual: Dict, funcionarios: Dict) -> pd.DataFrame:
        """
        Gera escala para uma semana específica usando RODÍZIO PERFEITO
        
        Args:
            semana_num: Número da semana (1-4)
            contadores: Dicionário com contadores históricos
            contadores_mes_atual: Contadores temporários do mês atual
            funcionarios: Dicionário com funcionários por ilha
            
        Returns:
            DataFrame com escala semanal
        """
        # Determinar qual ilha NÃO terá ninguém no domingo nesta semana
        ilha_sem_domingo = (semana_num - 1) % 4
        ilhas = list(funcionarios.keys())
        
        dados_semana = []
        
        for ilha_idx, ilha in enumerate(ilhas):
            lista_func = funcionarios[ilha]
            
            # DOMINGO: 1 pessoa por ilha (exceto a ilha do rodízio)
            trabalha_domingo = False
            funcionario_domingo = None
            
            if ilha_idx != ilha_sem_domingo:
                # Usar sistema de rodízio para escolher quem trabalha no domingo
                funcionario_domingo = self.obter_proximo_domingo(ilha)
                trabalha_domingo = True
                
                # Atualizar contador do mês atual
                contadores_mes_atual[funcionario_domingo]['domingos'] += 1
                contadores_mes_atual[funcionario_domingo]['total'] += 1
                
                print(f"    🏝️  {ilha}: {funcionario_domingo.split()[0]} no DOMINGO")
            
            # SÁBADO: 2 pessoas por ilha
            funcionarios_sabado = []
            for i in range(2):
                funcionario_sabado = self.obter_proximo_sabado(ilha)
                
                # Se a pessoa já foi escolhida para domingo, não pode fazer sábado
                while funcionario_sabado == funcionario_domingo:
                    # Tentar outro
                    funcionario_sabado = self.obter_proximo_sabado(ilha)
                
                funcionarios_sabado.append(funcionario_sabado)
                
                # Atualizar contador do mês atual
                contadores_mes_atual[funcionario_sabado]['sabados'] += 1
                contadores_mes_atual[funcionario_sabado]['total'] += 1
                
                if i == 0:
                    print(f"    🏝️  {ilha}: {funcionario_sabado.split()[0]} no SÁBADO")
            
            # Gerar escalas para todos os funcionários da ilha
            for funcionario in lista_func:
                # Verificar se trabalha no fim de semana
                trabalha_sabado = funcionario in funcionarios_sabado
                trabalha_domingo_func = (funcionario == funcionario_domingo)
                
                # Gerar escala do funcionário
                dias = self.gerar_escala_funcionario(
                    funcionario=funcionario,
                    trabalha_sabado=trabalha_sabado,
                    trabalha_domingo=trabalha_domingo_func,
                    ilha=ilha
                )
                
                # Calcular totais
                dias_trabalhados = dias.count('P')
                folgas = 7 - dias_trabalhados
                
                # Linha da escala
                linha = {
                    'Funcionário': funcionario,
                    'Ilha': ilha,
                    'Seg': dias[0],
                    'Ter': dias[1],
                    'Qua': dias[2],
                    'Qui': dias[3],
                    'Sex': dias[4],
                    'Sáb': dias[5],
                    'Dom': dias[6],
                    'Dias Trabalhados': dias_trabalhados,
                    'Folgas': folgas
                }
                
                dados_semana.append(linha)
        
        return pd.DataFrame(dados_semana)
    
    def gerar_escala_funcionario(self, funcionario: str, 
                                 trabalha_sabado: bool, 
                                 trabalha_domingo: bool,
                                 ilha: str) -> List[str]:
        """
        Gera escala individual para um funcionário com rodízio de folgas
        
        Args:
            funcionario: Nome do funcionário
            trabalha_sabado: Se trabalha no sábado
            trabalha_domingo: Se trabalha no domingo
            ilha: Nome da ilha
            
        Returns:
            Lista com 7 dias (P=Presente, F=Folga)
        """
        # REGRA 1: Não pode trabalhar sábado E domingo
        if trabalha_sabado and trabalha_domingo:
            # Prioridade ao DOMINGO (regra principal)
            trabalha_sabado = False
        
        # Começar com todos os dias como trabalho
        dias = ["P"] * 7
        
        # Marcar folgas no fim de semana se não trabalhar
        if not trabalha_sabado:
            dias[5] = "F"  # Sábado
        if not trabalha_domingo:
            dias[6] = "F"  # Domingo
        
        # Contar quantos dias de trabalho temos
        dias_trabalho = dias.count("P")
        
        # REGRA 2: Precisamos de 5 dias de trabalho
        if dias_trabalho > 5:
            # Temos dias extras para folgar
            dias_para_folgar = dias_trabalho - 5
            
            # Tentar folgar dias na semana (segunda a sexta)
            dias_semana_disponiveis = []
            for i in range(5):  # Segunda a sexta (índices 0-4)
                if dias[i] == "P":  # Apenas dias que estão marcados como trabalho
                    dias_semana_disponiveis.append(i)
            
            # Escolher os melhores dias para folgar baseado no rodízio
            dias_folga_escolhidos = []
            
            for _ in range(dias_para_folgar):
                if dias_semana_disponiveis:
                    melhor_dia = self.obter_melhor_folga_semanal(
                        ilha=ilha,
                        funcionario=funcionario,
                        dias_disponiveis=dias_semana_disponiveis
                    )
                    
                    dias_folga_escolhidos.append(melhor_dia)
                    dias_semana_disponiveis.remove(melhor_dia)
            
            # Aplicar as folgas escolhidas
            for dia in dias_folga_escolhidos:
                dias[dia] = "F"
        
        elif dias_trabalho < 5:
            # Precisamos de mais dias de trabalho
            dias_para_trabalhar = 5 - dias_trabalho
            
            # Encontrar dias que estão como folga (exceto fim de semana já definido)
            dias_folga_disponiveis = []
            for i in range(7):
                if dias[i] == "F":
                    # Não forçar trabalho no fim de semana se já foi definido como folga
                    if (i == 5 and not trabalha_sabado) or (i == 6 and not trabalha_domingo):
                        continue
                    dias_folga_disponiveis.append(i)
            
            # Escolher quais dias converter de folga para trabalho
            # Preferência: dias da semana primeiro
            dias_semana_folga = [d for d in dias_folga_disponiveis if d < 5]
            dias_fim_semana_folga = [d for d in dias_folga_disponiveis if d >= 5]
            
            dias_escolhidos = []
            
            # Primeiro tentar pegar dias da semana
            if len(dias_semana_folga) >= dias_para_trabalhar:
                # Escolher aleatoriamente entre dias da semana
                dias_escolhidos = random.sample(dias_semana_folga, dias_para_trabalhar)
            else:
                # Usar todos os dias da semana disponíveis
                dias_escolhidos = dias_semana_folga.copy()
                dias_restantes = dias_para_trabalhar - len(dias_semana_folga)
                
                # Se ainda precisar, usar fim de semana
                if dias_restantes > 0 and len(dias_fim_semana_folga) >= dias_restantes:
                    dias_escolhidos.extend(random.sample(dias_fim_semana_folga, dias_restantes))
            
            # Aplicar trabalho nos dias escolhidos
            for dia in dias_escolhidos:
                dias[dia] = "P"
        
        # REGRA 3: Não pode ter duas folgas seguidas na semana (segunda a sexta)
        # Verificar e corrigir se necessário
        dias_semana = dias[:5]  # Apenas segunda a sexta
        
        for i in range(len(dias_semana) - 1):
            if dias_semana[i] == "F" and dias_semana[i + 1] == "F":
                # Encontrar um dia de trabalho para trocar
                for j in range(len(dias_semana)):
                    if dias_semana[j] == "P" and abs(j - i) > 1:  # Evitar trocar com dia adjacente
                        # Trocar os dias
                        dias_semana[i], dias_semana[j] = dias_semana[j], dias_semana[i]
                        
                        # Atualizar o rodízio de folgas
                        self.rodizio_folgas[ilha]['contador_folgas'][funcionario][i] -= 1
                        self.rodizio_folgas[ilha]['contador_folgas'][funcionario][j] += 1
                        
                        # Atualizar histórico
                        if i in self.rodizio_folgas[ilha]['ultimas_folgas'][funcionario]:
                            self.rodizio_folgas[ilha]['ultimas_folgas'][funcionario].remove(i)
                        if j not in self.rodizio_folgas[ilha]['ultimas_folgas'][funcionario]:
                            self.rodizio_folgas[ilha]['ultimas_folgas'][funcionario].append(j)
                        
                        break
        
        # Atualizar os dias da semana
        for i in range(5):
            dias[i] = dias_semana[i]
        
        # REGRA 4: Garantir que temos exatamente 5 dias de trabalho
        while dias.count("P") > 5:
            indices_trabalho = [i for i, d in enumerate(dias) if d == "P"]
            
            # Preferir remover trabalho do fim de semana (se for extra)
            if len([i for i in indices_trabalho if i >= 5]) > (2 - dias[5:7].count("F")):
                # Tem trabalho extra no fim de semana
                for i in [5, 6]:  # Sábado e domingo
                    if i in indices_trabalho:
                        dias[i] = "F"
                        break
            else:
                # Remover do dia da semana com mais folgas históricas
                contadores = self.rodizio_folgas[ilha]['contador_folgas'][funcionario]
                dias_semana_trabalho = [i for i in indices_trabalho if i < 5]
                
                if dias_semana_trabalho:
                    # Encontrar o dia com mais folgas (menos problemático para folgar)
                    dia_para_folgar = max(dias_semana_trabalho, key=lambda d: contadores[d])
                    dias[dia_para_folgar] = "F"
                else:
                    # Se não há dias da semana, remover aleatoriamente
                    dias[random.choice(indices_trabalho)] = "F"
        
        while dias.count("P") < 5:
            indices_folga = [i for i, d in enumerate(dias) if d == "F"]
            
            if not indices_folga:
                break
                
            # Preferir adicionar trabalho em dias da semana
            dias_semana_folga = [i for i in indices_folga if i < 5]
            
            if dias_semana_folga:
                # Adicionar no dia com menos folgas históricas
                contadores = self.rodizio_folgas[ilha]['contador_folgas'][funcionario]
                dia_para_trabalhar = min(dias_semana_folga, key=lambda d: contadores[d])
                dias[dia_para_trabalhar] = "P"
            else:
                # Se só tem fim de semana, adicionar aleatoriamente
                dias[random.choice(indices_folga)] = "P"
        
        # VERIFICAÇÃO FINAL: Garantir que não trabalha sábado e domingo
        if dias[5] == "P" and dias[6] == "P":
            # Prioridade ao domingo
            dias[5] = "F"
            
            # Adicionar um dia de trabalho na semana
            for i in range(5):
                if dias[i] == "F":
                    dias[i] = "P"
                    break
        
        return dias
    
    def verificar_rodizio_perfeito(self, df_escala: pd.DataFrame) -> Dict:
        """
        Verifica se o rodízio perfeito foi seguido
        
        Args:
            df_escala: DataFrame com a escala
            
        Returns:
            Dicionário com resultados da verificação
        """
        resultados = {
            'rodizio_domingo_por_ilha': {},
            'rodizio_sabado_por_ilha': {},
            'violacoes_domingo': [],
            'violacoes_sabado': [],
            'violacoes_folgas': [],
            'balanceamento_perfeito': True
        }
        
        # Para cada ilha
        for ilha, lista_func in self.funcionarios.items():
            # Filtrar funcionários desta ilha
            df_ilha = df_escala[df_escala['Ilha'] == ilha]
            
            # Contar domingos por funcionário nesta escala
            domingos_por_func = {}
            for funcionario in lista_func:
                df_func = df_ilha[df_ilha['Funcionário'] == funcionario]
                domingos = (df_func['Dom'] == 'P').sum()
                domingos_por_func[funcionario] = domingos
            
            # Contar sábados por funcionário nesta escala
            sabados_por_func = {}
            for funcionario in lista_func:
                df_func = df_ilha[df_ilha['Funcionário'] == funcionario]
                sabados = (df_func['Sáb'] == 'P').sum()
                sabados_por_func[funcionario] = sabados
            
            resultados['rodizio_domingo_por_ilha'][ilha] = domingos_por_func
            resultados['rodizio_sabado_por_ilha'][ilha] = sabados_por_func
            
            # VERIFICAÇÃO 1: Ninguém pode ter mais de 1 domingo até que todos tenham pelo menos 1
            min_domingos = min(domingos_por_func.values())
            max_domingos = max(domingos_por_func.values())
            
            if max_domingos > min_domingos + 1:
                resultados['violacoes_domingo'].append(
                    f"{ilha}: Diferença muito grande em domingos ({min_domingos}-{max_domingos})"
                )
                resultados['balanceamento_perfeito'] = False
            
            # VERIFICAÇÃO 2: Ninguém pode ter mais de 2 sábados até que todos tenham pelo menos 1
            min_sabados = min(sabados_por_func.values())
            max_sabados = max(sabados_por_func.values())
            
            if max_sabados > min_sabados + 2:  # Mais flexível para sábado (2 pessoas por semana)
                resultados['violacoes_sabado'].append(
                    f"{ilha}: Diferença muito grande em sábados ({min_sabados}-{max_sabados})"
                )
                resultados['balanceamento_perfeito'] = False
        
        return resultados
    
    def salvar_escala_excel(self, df_escala: pd.DataFrame, ano: int, mes: int):
        """
        Salva a escala em um arquivo Excel com múltiplas abas
        
        Args:
            df_escala: DataFrame com a escala
            ano: Ano da escala
            mes: Mês da escala
        """
        # Nome do arquivo
        nome_arquivo = f"{self.diretorio_escalas}/ESCALA_{ano}_{mes:02d}.xlsx"
        
        # Calcular contadores totais
        contadores_totais = self.calcular_contadores(df_escala)
        
        # Calcular contadores acumulados (somando com mês anterior se existir)
        contadores_acumulados = self.calcular_contadores_acumulados(ano, mes, contadores_totais)
        
        # Verificar rodízio perfeito
        rodizio = self.verificar_rodizio_perfeito(df_escala)
        
        with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
            # ABA 1: ESCALA COMPLETA
            df_escala.to_excel(writer, sheet_name='ESCALA_COMPLETA', index=False)
            
            # ABA 2: RESUMO POR SEMANA
            self.criar_resumo_semanal(df_escala, writer)
            
            # ABA 3: RESUMO POR ILHA
            self.criar_resumo_ilha(df_escala, writer)
            
            # ABA 4: CONTADORES MÊS ATUAL
            df_contadores_mes = pd.DataFrame([
                {
                    'Funcionário': func,
                    'Sábados Trabalhados': cont['sabados'],
                    'Domingos Trabalhados': cont['domingos'],
                    'Total Fim de Semana': cont['total']
                }
                for func, cont in contadores_totais.items()
            ])
            df_contadores_mes.to_excel(writer, sheet_name='CONTADORES_MES_ATUAL', index=False)
            
            # ABA 5: CONTADORES ACUMULADOS COM RODADAS
            df_contadores_acum = pd.DataFrame([
                {
                    'Funcionário': func,
                    'Sábados Trabalhados': cont['sabados_trabalhados'],
                    'Domingos Trabalhados': cont['domingos_trabalhados'],
                    'Total Fim de Semana': cont['total_fim_semana'],
                    'Rodada Domingo': cont.get('rodada_domingo', 0),
                    'Rodada Sábado': cont.get('rodada_sabado', 0)
                }
                for func, cont in contadores_acumulados.items()
            ])
            df_contadores_acum.to_excel(writer, sheet_name='CONTADORES_FIM_SEMANA', index=False)
            
            # ABA 6: VERIFICAÇÃO DE REGRAS
            self.criar_verificacao_regras(df_escala, writer)
            
            # ABA 7: RODÍZIO PERFEITO (NOVA)
            self.criar_aba_rodizio_perfeito(rodizio, df_contadores_acum, writer)
            
            # ABA 8: DISTRIBUIÇÃO POR ILHA
            self.criar_aba_distribuicao_ilha(rodizio, writer)
            
            # ABA 9: RODÍZIO DE FOLGAS (NOVA)
            self.criar_aba_rodizio_folgas(writer)
        
        print(f"✅ Escala salva em: {nome_arquivo}")
        print(f"   - 9 abas incluídas no arquivo")
        
        # Mostrar resultados do rodízio
        print(f"\n📊 VERIFICAÇÃO DO RODÍZIO PERFEITO:")
        print(f"   Balanceamento perfeito: {'✅ SIM' if rodizio['balanceamento_perfeito'] else '❌ NÃO'}")
        
        if rodizio['violacoes_domingo']:
            print(f"\n⚠️  VIOLAÇÕES NO DOMINGO:")
            for violacao in rodizio['violacoes_domingo'][:3]:
                print(f"   • {violacao}")
        
        if rodizio['violacoes_sabado']:
            print(f"\n⚠️  VIOLAÇÕES NO SÁBADO:")
            for violacao in rodizio['violacoes_sabado'][:3]:
                print(f"   • {violacao}")
        
        return nome_arquivo
    
    def calcular_contadores(self, df_escala: pd.DataFrame) -> Dict:
        """
        Calcula contadores de fim de semana para o mês atual
        
        Args:
            df_escala: DataFrame com a escala
            
        Returns:
            Dicionário com contadores do mês
        """
        contadores = {}
        
        for funcionario in df_escala['Funcionário'].unique():
            df_func = df_escala[df_escala['Funcionário'] == funcionario]
            
            sabados = (df_func['Sáb'] == 'P').sum()
            domingos = (df_func['Dom'] == 'P').sum()
            total = sabados + domingos
            
            contadores[funcionario] = {
                'sabados': sabados,
                'domingos': domingos,
                'total': total
            }
        
        return contadores
    
    def calcular_contadores_acumulados(self, ano: int, mes: int, 
                                      contadores_mes_atual: Dict) -> Dict:
        """
        Calcula contadores acumulados (mês atual + meses anteriores)
        
        Args:
            ano: Ano atual
            mes: Mês atual
            contadores_mes_atual: Contadores do mês atual
            
        Returns:
            Dicionário com contadores acumulados
        """
        # Iniciar com contadores do mês atual
        contadores_acumulados = {}
        for func, cont in contadores_mes_atual.items():
            # Encontrar rodadas atuais para este funcionário
            rodada_domingo = 0
            rodada_sabado = 0
            
            for ilha, lista_func in self.funcionarios.items():
                if func in lista_func:
                    rodada_domingo = self.rodizio_ilhas[ilha]['rodada_domingo']
                    rodada_sabado = self.rodizio_ilhas[ilha]['rodada_sabado']
                    break
            
            contadores_acumulados[func] = {
                'sabados_trabalhados': cont['sabados'],
                'domingos_trabalhados': cont['domingos'],
                'total_fim_semana': cont['total'],
                'rodada_domingo': rodada_domingo,
                'rodada_sabado': rodada_sabado
            }
        
        # Tentar carregar contadores acumulados do mês anterior
        if mes == 1:
            ano_anterior = ano - 1
            mes_anterior = 12
        else:
            ano_anterior = ano
            mes_anterior = mes - 1
        
        arquivo_anterior = f"{self.diretorio_escalas}/ESCALA_{ano_anterior}_{mes_anterior:02d}.xlsx"
        
        if os.path.exists(arquivo_anterior):
            try:
                # Carregar contadores acumulados do mês anterior
                df_contadores_anterior = pd.read_excel(
                    arquivo_anterior, 
                    sheet_name='CONTADORES_FIM_SEMANA'
                )
                
                # Somar contadores
                for _, row in df_contadores_anterior.iterrows():
                    funcionario = row['Funcionário']
                    
                    if funcionario in contadores_acumulados:
                        # Somar com contadores do mês atual
                        contadores_acumulados[funcionario]['sabados_trabalhados'] += row['Sábados Trabalhados']
                        contadores_acumulados[funcionario]['domingos_trabalhados'] += row['Domingos Trabalhados']
                        contadores_acumulados[funcionario]['total_fim_semana'] += row['Total Fim de Semana']
                        
                        # Manter as rodadas mais recentes (do sistema atual)
                        # As rodadas são recalculadas a cada execução
                    else:
                        # Se funcionário não está no mês atual, usar apenas histórico
                        contadores_acumulados[funcionario] = {
                            'sabados_trabalhados': row['Sábados Trabalhados'],
                            'domingos_trabalhados': row['Domingos Trabalhados'],
                            'total_fim_semana': row['Total Fim de Semana'],
                            'rodada_domingo': row.get('Rodada Domingo', 0),
                            'rodada_sabado': row.get('Rodada Sábado', 0)
                        }
                
                print(f"✅ Contadores acumulados com mês anterior: {mes_anterior:02d}/{ano_anterior}")
                
            except Exception as e:
                print(f"⚠️  Não foi possível carregar contadores acumulados: {e}")
        
        return contadores_acumulados
    
    def criar_resumo_semanal(self, df_escala: pd.DataFrame, writer):
        """Cria aba de resumo semanal"""
        resumo = df_escala.groupby(['Semana do Mês']).agg({
            'Funcionário': 'count',
            'Dias Trabalhados': 'sum',
            'Sáb': lambda x: (x == 'P').sum(),
            'Dom': lambda x: (x == 'P').sum()
        }).reset_index()
        
        resumo.columns = ['Semana', 'Total Funcionários', 'Total Dias Trabalhados',
                         'Pessoas no Sábado', 'Pessoas no Domingo']
        
        resumo['Meta Sábado'] = 8
        resumo['Meta Domingo'] = 3
        resumo['Status Sábado'] = resumo.apply(
            lambda x: '✅ OK' if x['Pessoas no Sábado'] == 8 else f'❌ Faltam {8 - x["Pessoas no Sábado"]}', 
            axis=1
        )
        resumo['Status Domingo'] = resumo.apply(
            lambda x: '✅ OK' if x['Pessoas no Domingo'] == 3 else f'❌ Faltam {3 - x["Pessoas no Domingo"]}', 
            axis=1
        )
        
        resumo.to_excel(writer, sheet_name='RESUMO_SEMANAL', index=False)
    
    def criar_resumo_ilha(self, df_escala: pd.DataFrame, writer):
        """Cria aba de resumo por ilha"""
        resumo_ilha = df_escala.groupby(['Ilha', 'Semana do Mês']).agg({
            'Funcionário': 'count',
            'Sáb': lambda x: (x == 'P').sum(),
            'Dom': lambda x: (x == 'P').sum()
        }).reset_index()
        
        resumo_ilha.columns = ['Ilha', 'Semana', 'Total Funcionários',
                              'Sábados Trabalhados', 'Domingos Trabalhados']
        
        resumo_ilha.to_excel(writer, sheet_name='RESUMO_POR_ILHA', index=False)
    
    def criar_verificacao_regras(self, df_escala: pd.DataFrame, writer):
        """Cria aba de verificação de regras"""
        verificacao = self.verificar_regras(df_escala)
        
        dados_verificacao = [
            ['Regra', 'Status'],
            ['5 dias trabalhados por semana', '✅ OK' if verificacao['regra_5_dias'] else '❌ FALHOU'],
            ['Sem duas folgas seguidas (seg-sex)', '✅ OK' if verificacao['regra_folgas_seguidas'] else '❌ FALHOU'],
            ['Não trabalha sábado e domingo', '✅ OK' if verificacao['regra_fim_semana_seguido'] else '❌ FALHOU'],
            ['Cobertura de sábado (8 pessoas)', '✅ OK' if verificacao['cobertura_sabado'] else '❌ FALHOU'],
            ['Cobertura de domingo (3 pessoas)', '✅ OK' if verificacao['cobertura_domingo'] else '❌ FALHOU'],
            ['Rodízio de domingo perfeito', '✅ OK' if verificacao['rodizio_domingo'] else '❌ FALHOU'],
            ['Rodízio de sábado perfeito', '✅ OK' if verificacao['rodizio_sabado'] else '❌ FALHOU'],
            ['Rodízio de folgas semanal', '✅ OK' if verificacao['rodizio_folgas'] else '❌ FALHOU'],
        ]
        
        df_verificacao = pd.DataFrame(dados_verificacao[1:], columns=dados_verificacao[0])
        
        # Adicionar erros específicos se houver
        if verificacao['erros']:
            df_erros = pd.DataFrame({'Erros Detectados': verificacao['erros'][:10]})
            df_erros.to_excel(writer, sheet_name='ERROS_DETECTADOS', index=False)
        
        df_verificacao.to_excel(writer, sheet_name='VERIFICACAO_REGRAS', index=False)
    
    def criar_aba_rodizio_perfeito(self, rodizio: Dict, df_contadores: pd.DataFrame, writer):
        """Cria nova aba de rodízio perfeito"""
        # Adicionar informações de rodízio
        df_rodizio = df_contadores.copy()
        
        # Calcular diferença dentro de cada ilha
        def calcular_diferenca_ilha(funcionario):
            for ilha, lista_func in self.funcionarios.items():
                if funcionario in lista_func:
                    # Encontrar todos da mesma ilha
                    mesma_ilha = [f for f in df_rodizio['Funcionário'] if f in lista_func]
                    df_ilha = df_rodizio[df_rodizio['Funcionário'].isin(mesma_ilha)]
                    
                    min_dom = df_ilha['Domingos Trabalhados'].min()
                    max_dom = df_ilha['Domingos Trabalhados'].max()
                    min_sab = df_ilha['Sábados Trabalhados'].min()
                    max_sab = df_ilha['Sábados Trabalhados'].max()
                    
                    diff_dom = max_dom - min_dom
                    diff_sab = max_sab - min_sab
                    
                    status_dom = "✅" if diff_dom <= 1 else "⚠️"
                    status_sab = "✅" if diff_sab <= 1 else "⚠️"
                    
                    return f"Dom: {min_dom}-{max_dom} {status_dom} | Sáb: {min_sab}-{max_sab} {status_sab}"
            return "N/A"
        
        df_rodizio['Diferença na Ilha'] = df_rodizio['Funcionário'].apply(calcular_diferenca_ilha)
        
        # Ordenar por menos domingos
        df_rodizio = df_rodizio.sort_values(['Domingos Trabalhados', 'Sábados Trabalhados'], 
                                          ascending=[True, True])
        
        df_rodizio.to_excel(writer, sheet_name='RODÍZIO_PERFEITO', index=False)
    
    def criar_aba_distribuicao_ilha(self, rodizio: Dict, writer):
        """Cria aba com distribuição detalhada por ilha"""
        dados_ilha = []
        
        for ilha in self.funcionarios.keys():
            if ilha in rodizio['rodizio_domingo_por_ilha']:
                domingos = rodizio['rodizio_domingo_por_ilha'][ilha]
                sabados = rodizio['rodizio_sabado_por_ilha'][ilha]
                
                for funcionario in domingos.keys():
                    dados_ilha.append({
                        'Ilha': ilha,
                        'Funcionário': funcionario,
                        'Domingos (mês)': domingos[funcionario],
                        'Sábados (mês)': sabados[funcionario],
                        'Total Fim de Semana (mês)': domingos[funcionario] + sabados[funcionario]
                    })
        
        if dados_ilha:
            df_distribuicao = pd.DataFrame(dados_ilha)
            df_distribuicao.to_excel(writer, sheet_name='DISTRIBUIÇÃO_ILHA', index=False)
    
    def criar_aba_rodizio_folgas(self, writer):
        """Cria nova aba de rodízio de folgas"""
        dados_folgas = []
        
        for ilha in self.rodizio_folgas.keys():
            rodizio = self.rodizio_folgas[ilha]
            
            for funcionario in rodizio['contador_folgas'].keys():
                contadores = rodizio['contador_folgas'][funcionario]
                
                dados_folgas.append({
                    'Ilha': ilha,
                    'Funcionário': funcionario,
                    'Folgas Segunda': contadores[0],
                    'Folgas Terça': contadores[1],
                    'Folgas Quarta': contadores[2],
                    'Folgas Quinta': contadores[3],
                    'Folgas Sexta': contadores[4],
                    'Total Folgas Semana': sum(contadores.values()),
                    'Últimas Folgas': ', '.join([self.dias_completos[i] for i in rodizio['ultimas_folgas'][funcionario][-3:]]) if rodizio['ultimas_folgas'][funcionario] else 'Nenhuma'
                })
        
        if dados_folgas:
            df_folgas = pd.DataFrame(dados_folgas)
            
            # Calcular estatísticas por ilha
            estatisticas_ilha = []
            for ilha in self.rodizio_folgas.keys():
                df_ilha = df_folgas[df_folgas['Ilha'] == ilha]
                
                for dia_idx, dia_nome in enumerate(['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']):
                    media = df_ilha[f'Folgas {dia_nome}'].mean()
                    min_val = df_ilha[f'Folgas {dia_nome}'].min()
                    max_val = df_ilha[f'Folgas {dia_nome}'].max()
                    diff = max_val - min_val
                    
                    estatisticas_ilha.append({
                        'Ilha': ilha,
                        'Dia da Semana': dia_nome,
                        'Média de Folgas': round(media, 2),
                        'Mínimo': min_val,
                        'Máximo': max_val,
                        'Diferença': diff,
                        'Status': '✅ Balanceado' if diff <= 2 else '⚠️ Desbalanceado'
                    })
            
            df_estatisticas = pd.DataFrame(estatisticas_ilha)
            
            # Ordenar por total de folgas
            df_folgas = df_folgas.sort_values('Total Folgas Semana', ascending=True)
            
            # Salvar em abas separadas
            df_folgas.to_excel(writer, sheet_name='RODÍZIO_FOLGAS', index=False)
            df_estatisticas.to_excel(writer, sheet_name='ESTAT_FOLGAS', index=False)
    
    def verificar_regras(self, df_escala: pd.DataFrame) -> Dict:
        """
        Verifica se todas as regras foram atendidas
        
        Args:
            df_escala: DataFrame com a escala
            
        Returns:
            Dicionário com resultados da verificação
        """
        resultados = {
            'regra_5_dias': True,
            'regra_folgas_seguidas': True,
            'regra_fim_semana_seguido': True,
            'cobertura_sabado': True,
            'cobertura_domingo': True,
            'rodizio_domingo': True,
            'rodizio_sabado': True,
            'rodizio_folgas': True,
            'erros': []
        }
        
        # REGRA 1: Cada funcionário deve ter 5 dias de trabalho
        for _, row in df_escala.iterrows():
            if row['Dias Trabalhados'] != 5:
                resultados['regra_5_dias'] = False
                resultados['erros'].append(
                    f"{row['Funcionário']} tem {row['Dias Trabalhados']} dias trabalhados"
                )
        
        # REGRA 2: Não pode ter duas folgas seguidas na semana
        for _, row in df_escala.iterrows():
            dias_semana = [row['Seg'], row['Ter'], row['Qua'], row['Qui'], row['Sex']]
            
            for i in range(len(dias_semana) - 1):
                if dias_semana[i] == "F" and dias_semana[i + 1] == "F":
                    resultados['regra_folgas_seguidas'] = False
                    resultados['erros'].append(
                        f"{row['Funcionário']} tem folgas seguidas: {self.dias_completos[i]} e {self.dias_completos[i+1]}"
                    )
                    break
        
        # REGRA 3: Não pode trabalhar sábado e domingo
        for _, row in df_escala.iterrows():
            if row['Sáb'] == "P" and row['Dom'] == "P":
                resultados['regra_fim_semana_seguido'] = False
                resultados['erros'].append(
                    f"{row['Funcionário']} trabalha sábado e domingo"
                )
        
        # REGRA 4: Cobertura de sábado (8 pessoas)
        semanas = df_escala['Semana do Mês'].unique()
        for semana in semanas:
            df_semana = df_escala[df_escala['Semana do Mês'] == semana]
            pessoas_sabado = (df_semana['Sáb'] == 'P').sum()
            
            if pessoas_sabado != 8:
                resultados['cobertura_sabado'] = False
                resultados['erros'].append(
                    f"Semana {semana}: {pessoas_sabado} pessoas no sábado (deveria ser 8)"
                )
        
        # REGRA 5: Cobertura de domingo (3 pessoas)
        for semana in semanas:
            df_semana = df_escala[df_escala['Semana do Mês'] == semana]
            pessoas_domingo = (df_semana['Dom'] == 'P').sum()
            
            if pessoas_domingo != 3:
                resultados['cobertura_domingo'] = False
                resultados['erros'].append(
                    f"Semana {semana}: {pessoas_domingo} pessoas no domingo (deveria ser 3)"
                )
        
        # REGRA 6: Rodízio de domingo
        rodizio = self.verificar_rodizio_perfeito(df_escala)
        if not rodizio['balanceamento_perfeito']:
            resultados['rodizio_domingo'] = False
            resultados['rodizio_sabado'] = False
            for violacao in rodizio['violacoes_domingo'][:2]:
                resultados['erros'].append(f"Rodízio Domingo: {violacao}")
            for violacao in rodizio['violacoes_sabado'][:2]:
                resultados['erros'].append(f"Rodízio Sábado: {violacao}")
        
        # REGRA 7: Rodízio de folgas (verificar se as folgas estão bem distribuídas)
        # Analisar por ilha
        for ilha in self.funcionarios.keys():
            df_ilha = df_escala[df_escala['Ilha'] == ilha]
            
            # Contar folgas por dia da semana
            folgas_por_dia = {i: 0 for i in range(5)}
            for _, row in df_ilha.iterrows():
                for i, dia in enumerate(['Seg', 'Ter', 'Qua', 'Qui', 'Sex']):
                    if row[dia] == 'F':
                        folgas_por_dia[i] += 1
            
            # Verificar se as folgas estão bem distribuídas
            total_funcionarios = len(df_ilha['Funcionário'].unique())
            semanas_no_mes = len(semanas)
            
            # Idealmente, cada funcionário deveria ter folgas em dias diferentes
            # Verificar se há dias com muitas ou poucas folgas
            for dia_idx, dia_nome in enumerate(['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']):
                folgas_no_dia = folgas_por_dia[dia_idx]
                media_esperada = (total_funcionarios * 2 * semanas_no_mes) / 5  # 2 folgas por semana por funcionário
                
                if abs(folgas_no_dia - media_esperada) > media_esperada * 0.5:  # 50% de tolerância
                    resultados['rodizio_folgas'] = False
                    resultados['erros'].append(
                        f"{ilha}: Folgas na {dia_nome} desbalanceadas ({folgas_no_dia} vs esperado ~{media_esperada:.1f})"
                    )
                    break
        
        return resultados
    
    def listar_escalas_existentes(self, return_list=False):
        """Lista todas as escalas existentes no diretório"""
        arquivos = sorted([f for f in os.listdir(self.diretorio_escalas) 
                          if f.startswith('ESCALA_') and f.endswith('.xlsx')])
        
        if return_list:
            return arquivos
        
        print("\n" + "=" * 60)
        print("ESCALAS EXISTENTES NO HISTÓRICO")
        print("=" * 60)
        
        if not arquivos:
            print("📭 Nenhuma escala encontrada no histórico.")
            return []
        
        print(f"📁 Total de escalas: {len(arquivos)}")
        print("\nArquivos encontrados:")
        for i, arquivo in enumerate(arquivos, 1):
            # Extrair ano e mês do nome do arquivo
            partes = arquivo.replace('ESCALA_', '').replace('.xlsx', '').split('_')
            if len(partes) >= 2:
                ano, mes = partes[0], partes[1]
                print(f"  {i:2d}. {arquivo} → {mes}/{ano}")
        
        return arquivos
    
    def carregar_escala_anterior(self, ano: int, mes: int) -> Optional[pd.DataFrame]:
        """
        Carrega a escala do mês anterior
        
        Args:
            ano: Ano atual
            mes: Mês atual
            
        Returns:
            DataFrame da escala anterior ou None
        """
        if mes == 1:
            ano_anterior = ano - 1
            mes_anterior = 12
        else:
            ano_anterior = ano
            mes_anterior = mes - 1
        
        arquivo = f"{self.diretorio_escalas}/ESCALA_{ano_anterior}_{mes_anterior:02d}.xlsx"
        
        if os.path.exists(arquivo):
            try:
                df_anterior = pd.read_excel(arquivo, sheet_name='ESCALA_COMPLETA')
                print(f"✅ Escala anterior carregada: {mes_anterior:02d}/{ano_anterior}")
                return df_anterior
            except Exception as e:
                print(f"❌ Erro ao carregar escala anterior: {e}")
        
        return None
    
    def gerar_relatorio_anual(self, ano: int):
        """
        Gera um relatório consolidado de um ano inteiro
        
        Args:
            ano: Ano do relatório
        """
        print(f"\n📊 GERANDO RELATÓRIO ANUAL {ano}")
        print("=" * 50)
        
        # Encontrar todas as escalas do ano
        escalas_ano = []
        
        for mes in range(1, 13):
            arquivo = f"{self.diretorio_escalas}/ESCALA_{ano}_{mes:02d}.xlsx"
            if os.path.exists(arquivo):
                try:
                    df_mes = pd.read_excel(arquivo, sheet_name='ESCALA_COMPLETA')
                    escalas_ano.append(df_mes)
                    print(f"  ✓ Mês {mes:02d}: {len(df_mes)} registros")
                except:
                    print(f"  ✗ Mês {mes:02d}: erro ao carregar")
        
        if not escalas_ano:
            print(f"❌ Nenhuma escala encontrada para o ano {ano}")
            return
        
        # Consolidar todas as escalas
        df_anual = pd.concat(escalas_ano, ignore_index=True)
        
        # Criar diretório para relatórios anuais
        dir_relatorios = "RELATORIOS_ANUAIS"
        os.makedirs(dir_relatorios, exist_ok=True)
        
        # Nome do arquivo do relatório anual
        arquivo_relatorio = f"{dir_relatorios}/RELATORIO_ANUAL_{ano}.xlsx"
        
        with pd.ExcelWriter(arquivo_relatorio, engine='openpyxl') as writer:
            # ABA 1: DADOS COMPLETOS DO ANO
            df_anual.to_excel(writer, sheet_name='DADOS_ANUAIS', index=False)
            
            # ABA 2: ESTATÍSTICAS POR FUNCIONÁRIO
            stats_func = df_anual.groupby(['Funcionário', 'Ilha']).agg({
                'Dias Trabalhados': 'sum',
                'Sáb': lambda x: (x == 'P').sum(),
                'Dom': lambda x: (x == 'P').sum()
            }).reset_index()
            
            stats_func['Total Fim de Semana'] = stats_func['Sáb'] + stats_func['Dom']
            stats_func.to_excel(writer, sheet_name='ESTATISTICAS_FUNCIONARIOS', index=False)
            
            # ABA 3: ESTATÍSTICAS POR MÊS
            stats_mes = df_anual.groupby(['Mês']).agg({
                'Funcionário': 'nunique',
                'Dias Trabalhados': 'sum',
                'Sáb': lambda x: (x == 'P').sum(),
                'Dom': lambda x: (x == 'P').sum()
            }).reset_index()
            
            stats_mes.columns = ['Mês', 'Funcionários Únicos', 'Total Dias Trabalhados',
                                'Sábados Trabalhados', 'Domingos Trabalhados']
            stats_mes.to_excel(writer, sheet_name='ESTATISTICAS_MENSAL', index=False)
            
            # ABA 4: BALANCEAMENTO DE FIM DE SEMANA
            balanceamento = stats_func.copy()
            balanceamento['Média Mensal'] = balanceamento['Total Fim de Semana'] / len(df_anual['Mês'].unique())
            balanceamento = balanceamento.sort_values('Total Fim de Semana', ascending=True)
            balanceamento.to_excel(writer, sheet_name='BALANCEAMENTO_ANUAL', index=False)
            
            # ABA 5: RESUMO GERAL
            resumo_geral = pd.DataFrame({
                'Métrica': [
                    'Total de Meses',
                    'Total de Funcionários',
                    'Total de Dias Trabalhados',
                    'Média Dias/Funcionário/Mês',
                    'Total Sábados Trabalhados',
                    'Total Domingos Trabalhados',
                    'Média Sábados/Funcionário',
                    'Média Domingos/Funcionário',
                    'Funcionários sem Domingo',
                    'Funcionários com 1 Domingo',
                    'Funcionários com 2+ Domingos',
                    'Maior diferença em Domingos',
                    'Maior diferença em Sábados'
                ],
                'Valor': [
                    len(df_anual['Mês'].unique()),
                    len(df_anual['Funcionário'].unique()),
                    df_anual['Dias Trabalhados'].sum(),
                    df_anual['Dias Trabalhados'].sum() / (len(df_anual['Funcionário'].unique()) * len(df_anual['Mês'].unique())),
                    stats_mes['Sábados Trabalhados'].sum(),
                    stats_mes['Domingos Trabalhados'].sum(),
                    stats_mes['Sábados Trabalhados'].sum() / len(df_anual['Funcionário'].unique()),
                    stats_mes['Domingos Trabalhados'].sum() / len(df_anual['Funcionário'].unique()),
                    (stats_func['Dom'] == 0).sum(),
                    (stats_func['Dom'] == 1).sum(),
                    (stats_func['Dom'] >= 2).sum(),
                    stats_func['Dom'].max() - stats_func['Dom'].min() if len(stats_func) > 0 else 0,
                    stats_func['Sáb'].max() - stats_func['Sáb'].min() if len(stats_func) > 0 else 0
                ]
            })
            resumo_geral.to_excel(writer, sheet_name='RESUMO_GERAL', index=False)
        
        print(f"\n✅ Relatório anual {ano} salvo em: {arquivo_relatorio}")
        print(f"   - 5 abas incluídas no relatório")
        
        return arquivo_relatorio