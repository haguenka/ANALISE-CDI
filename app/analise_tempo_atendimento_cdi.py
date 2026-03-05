#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Centro de Diagnóstico por Imagem - Análise de Tempo de Atendimento
Casa de Saúde São José

Análise de tempo de atendimento para exames do CDI
Tempo máximo esperado: 1 hora
"""

import sys
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path
import json
import os

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                               QHBoxLayout, QLayout, QPushButton, QLabel, QFileDialog,
                               QDateEdit, QGroupBox, QGridLayout, QScrollArea,
                               QComboBox, QTableWidget, QTableWidgetItem, QHeaderView,
                               QFrame, QSplitter, QMessageBox, QTextEdit, QLineEdit,
                               QProgressBar, QTabWidget, QMenuBar, QWidgetAction, QSpinBox,
                               QGraphicsDropShadowEffect)
from PySide6.QtCore import Qt, QDate, Signal, QThread
from PySide6.QtGui import QFont, QPalette, QColor, QIcon
from PySide6.QtCharts import QChart, QChartView, QPieSeries, QBarSeries, QBarSet, QBarCategoryAxis, QValueAxis

import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

# AI Integration
try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False


def get_runtime_data_dir():
    """Retorna um diretorio valido para abrir/salvar arquivos no ambiente atual."""
    preferred = Path(os.getenv("CDI_DATA_DIR", "/var/data"))
    if preferred.exists():
        return preferred

    try:
        preferred.mkdir(parents=True, exist_ok=True)
        return preferred
    except Exception:
        fallback = Path.home() / "cdi_data"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback


def build_runtime_file_path(filename):
    """Monta um caminho padrao dentro do diretorio de runtime."""
    return str(get_runtime_data_dir() / filename)


class PatientSearchWindow(QMainWindow):
    """Janela de busca de paciente com histórico de exames"""
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.df = df
        self.setWindowTitle("🔍 Buscar Paciente - CDI São José")
        self.setGeometry(100, 100, 1200, 800)

        # Aplicar dark theme
        self.setStyleSheet("""
            QMainWindow {
                background-color: #0d1117;
            }
            QWidget {
                background-color: #0d1117;
                color: #c9d1d9;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
            }
            QLabel {
                color: #c9d1d9;
            }
            QLineEdit {
                background-color: #21262d;
                border: 1px solid #30363d;
                border-radius: 6px;
                padding: 8px;
                color: #c9d1d9;
                font-size: 14px;
            }
            QLineEdit:focus {
                border-color: #58a6ff;
            }
            QPushButton {
                background-color: #21262d;
                border: 1px solid #30363d;
                border-radius: 6px;
                padding: 8px 16px;
                color: #c9d1d9;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #30363d;
                border-color: #58a6ff;
            }
            QPushButton#searchButton {
                background-color: #238636;
                border-color: #238636;
            }
            QPushButton#searchButton:hover {
                background-color: #2ea043;
            }
            QPushButton#exportButton {
                background-color: #1f6feb;
                border-color: #1f6feb;
            }
            QPushButton#exportButton:hover {
                background-color: #388bfd;
            }
            QTableWidget {
                background-color: #0d1117;
                alternate-background-color: #161b22;
                border: 1px solid #30363d;
                border-radius: 6px;
                gridline-color: #30363d;
            }
            QTableWidget::item {
                padding: 5px;
                color: #c9d1d9;
            }
            QTableWidget::item:selected {
                background-color: #58a6ff;
                color: #ffffff;
            }
            QHeaderView::section {
                background-color: #161b22;
                color: #58a6ff;
                padding: 8px;
                border: none;
                border-right: 1px solid #30363d;
                border-bottom: 1px solid #30363d;
                font-weight: bold;
            }
            QGroupBox {
                background-color: #161b22;
                border: 1px solid #30363d;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 15px;
                font-weight: bold;
            }
            QGroupBox::title {
                color: #58a6ff;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)

        self.setup_ui()

    def setup_ui(self):
        """Configura a interface"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Header
        header = QLabel("🔍 Buscar Paciente - Centro de Diagnóstico por Imagem")
        header.setStyleSheet("""
            font-size: 20px;
            font-weight: bold;
            color: #58a6ff;
            padding: 10px;
        """)
        main_layout.addWidget(header)

        # Busca
        search_group = QGroupBox("Buscar Paciente")
        search_layout = QHBoxLayout()

        search_label = QLabel("Digite SAME ou Nome:")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Ex: 12345 ou MARIA SILVA")
        self.search_input.returnPressed.connect(self.search_patient)

        self.search_btn = QPushButton("🔍 Buscar")
        self.search_btn.setObjectName("searchButton")
        self.search_btn.clicked.connect(self.search_patient)

        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input, stretch=1)
        search_layout.addWidget(self.search_btn)

        search_group.setLayout(search_layout)
        main_layout.addWidget(search_group)

        # Dados do paciente
        self.patient_info_group = QGroupBox("Dados do Paciente")
        patient_info_layout = QGridLayout()

        self.same_label = QLabel("SAME:")
        self.same_value = QLabel("-")
        self.same_value.setStyleSheet("font-size: 14px; font-weight: bold; color: #58a6ff;")

        self.nome_label = QLabel("Nome:")
        self.nome_value = QLabel("-")
        self.nome_value.setStyleSheet("font-size: 14px; font-weight: bold; color: #58a6ff;")

        self.nasc_label = QLabel("Data Nascimento:")
        self.nasc_value = QLabel("-")

        self.idade_label = QLabel("Idade:")
        self.idade_value = QLabel("-")

        patient_info_layout.addWidget(self.same_label, 0, 0)
        patient_info_layout.addWidget(self.same_value, 0, 1)
        patient_info_layout.addWidget(self.nome_label, 0, 2)
        patient_info_layout.addWidget(self.nome_value, 0, 3)
        patient_info_layout.addWidget(self.nasc_label, 1, 0)
        patient_info_layout.addWidget(self.nasc_value, 1, 1)
        patient_info_layout.addWidget(self.idade_label, 1, 2)
        patient_info_layout.addWidget(self.idade_value, 1, 3)

        self.patient_info_group.setLayout(patient_info_layout)
        self.patient_info_group.setVisible(False)
        main_layout.addWidget(self.patient_info_group)

        # Histórico de exames
        self.exams_group = QGroupBox("Histórico de Exames")
        exams_layout = QVBoxLayout()

        # Tabela
        self.exams_table = QTableWidget()
        self.exams_table.setColumnCount(11)
        self.exams_table.setHorizontalHeaderLabels([
            'Data/Hora Prescrição',
            'Grupo',
            'Tipo Atendimento',
            'Data/Hora Laudo',
            'Tempo Realização (min)',
            'SLA Realização',
            'Data/Hora Entrega',
            'Tempo Liberação (min)',
            'SLA Entrega',
            'SLA Esperado',
            'Dias Úteis'
        ])

        # Configurar tabela
        self.exams_table.setAlternatingRowColors(True)
        self.exams_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.exams_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.exams_table.horizontalHeader().setStretchLastSection(True)
        self.exams_table.verticalHeader().setVisible(False)

        # Ajustar largura das colunas
        header = self.exams_table.horizontalHeader()
        for i in range(11):
            if i in [0, 3, 6]:  # Colunas de data
                header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
            elif i in [1, 2, 9]:  # Grupo, Tipo, SLA Esperado
                header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
            else:
                header.setSectionResizeMode(i, QHeaderView.ResizeToContents)

        exams_layout.addWidget(self.exams_table)

        # Botão exportar
        export_layout = QHBoxLayout()
        export_layout.addStretch()
        self.export_btn = QPushButton("📊 Exportar para Excel")
        self.export_btn.setObjectName("exportButton")
        self.export_btn.clicked.connect(self.export_to_excel)
        self.export_btn.setEnabled(False)
        export_layout.addWidget(self.export_btn)

        exams_layout.addLayout(export_layout)

        self.exams_group.setLayout(exams_layout)
        self.exams_group.setVisible(False)
        main_layout.addWidget(self.exams_group, stretch=1)

    def search_patient(self):
        """Busca paciente por SAME ou Nome"""
        search_term = self.search_input.text().strip()

        if not search_term:
            QMessageBox.warning(self, "Aviso", "Por favor, digite um SAME ou Nome para buscar.")
            return

        # Verificar se é busca numérica (SAME) ou texto (NOME)
        if search_term.isdigit():
            # Busca por SAME
            results = self.df[self.df['SAME'].astype(str).str.contains(search_term, na=False)]
        else:
            # Busca por nome (case-insensitive)
            results = self.df[self.df['NOME_PACIENTE'].str.contains(search_term, case=False, na=False)]

        if len(results) == 0:
            QMessageBox.information(self, "Sem resultados",
                                   f"Nenhum paciente encontrado com o termo: {search_term}")
            self.patient_info_group.setVisible(False)
            self.exams_group.setVisible(False)
            return

        # Se encontrou múltiplos pacientes com nomes diferentes, mostrar aviso
        unique_patients = results['SAME'].unique()
        if len(unique_patients) > 1:
            patient_list = []
            for same in unique_patients[:10]:  # Limitar a 10 para não ficar muito grande
                patient = results[results['SAME'] == same].iloc[0]
                patient_list.append(f"SAME {int(same)}: {patient['NOME_PACIENTE']}")

            msg = f"Encontrados {len(unique_patients)} pacientes:\n\n" + "\n".join(patient_list)
            if len(unique_patients) > 10:
                msg += f"\n\n... e mais {len(unique_patients) - 10} pacientes"
            msg += "\n\nPor favor, refine sua busca digitando um SAME específico."

            QMessageBox.information(self, "Múltiplos pacientes", msg)
            return

        # Mostrar dados do paciente
        self.display_patient_data(results)

    def display_patient_data(self, patient_exams):
        """Mostra dados do paciente e histórico de exames"""
        # Pegar primeira linha para dados do paciente
        first_row = patient_exams.iloc[0]

        # Atualizar dados do paciente
        same = int(first_row['SAME']) if pd.notna(first_row['SAME']) else 'N/A'
        self.same_value.setText(str(same))
        self.nome_value.setText(str(first_row['NOME_PACIENTE']) if pd.notna(first_row['NOME_PACIENTE']) else 'N/A')

        # Data de nascimento
        if 'DATA_NASCIMENTO' in patient_exams.columns and pd.notna(first_row['DATA_NASCIMENTO']):
            if isinstance(first_row['DATA_NASCIMENTO'], str):
                self.nasc_value.setText(first_row['DATA_NASCIMENTO'])
            else:
                self.nasc_value.setText(first_row['DATA_NASCIMENTO'].strftime('%d/%m/%Y'))
        else:
            self.nasc_value.setText('N/A')

        # Idade
        if 'IDADE' in patient_exams.columns and pd.notna(first_row['IDADE']):
            self.idade_value.setText(f"{int(first_row['IDADE'])} anos")
        else:
            self.idade_value.setText('N/A')

        self.patient_info_group.setVisible(True)

        # Preencher tabela de exames
        self.populate_exams_table(patient_exams)
        self.exams_group.setVisible(True)
        self.export_btn.setEnabled(True)

    def populate_exams_table(self, patient_exams):
        """Preenche tabela com histórico de exames"""
        # Ordenar por data de prescrição (mais recente primeiro)
        patient_exams = patient_exams.sort_values('DATA_HORA_PRESCRICAO', ascending=False)

        self.exams_table.setRowCount(len(patient_exams))

        for i, (idx, row) in enumerate(patient_exams.iterrows()):
            # Data/Hora Prescrição
            if pd.notna(row['DATA_HORA_PRESCRICAO']):
                prescricao_item = QTableWidgetItem(row['DATA_HORA_PRESCRICAO'].strftime('%d/%m/%Y %H:%M'))
            else:
                prescricao_item = QTableWidgetItem('N/A')
            self.exams_table.setItem(i, 0, prescricao_item)

            # Grupo
            grupo_item = QTableWidgetItem(str(row['GRUPO']) if pd.notna(row['GRUPO']) else 'N/A')
            self.exams_table.setItem(i, 1, grupo_item)

            # Tipo Atendimento
            tipo_item = QTableWidgetItem(str(row['TIPO_ATENDIMENTO']) if pd.notna(row['TIPO_ATENDIMENTO']) else 'N/A')
            self.exams_table.setItem(i, 2, tipo_item)

            # Data/Hora Laudo
            if pd.notna(row['STATUS_ALAUDAR']):
                laudo_item = QTableWidgetItem(row['STATUS_ALAUDAR'].strftime('%d/%m/%Y %H:%M'))
            else:
                laudo_item = QTableWidgetItem('N/A')
            self.exams_table.setItem(i, 3, laudo_item)

            # Tempo Realização (DATA_HORA_PRESCRICAO até STATUS_ALAUDAR)
            if 'TEMPO_ATENDIMENTO_MIN' in row and pd.notna(row['TEMPO_ATENDIMENTO_MIN']):
                tempo_real = row['TEMPO_ATENDIMENTO_MIN']
                if tempo_real > 1440:  # Mais de 24 horas
                    dias = tempo_real / 1440
                    tempo_text = f"{tempo_real:.1f} ({dias:.1f} dias)"
                else:
                    tempo_text = f"{tempo_real:.1f}"
                tempo_item = QTableWidgetItem(tempo_text)

                # Colorir baseado no tempo
                if tempo_real > 120:
                    tempo_item.setForeground(QColor('#F44336'))  # Vermelho
                elif tempo_real > 60:
                    tempo_item.setForeground(QColor('#FF9800'))  # Laranja
                else:
                    tempo_item.setForeground(QColor('#4CAF50'))  # Verde
            else:
                tempo_item = QTableWidgetItem('N/A')
            self.exams_table.setItem(i, 4, tempo_item)

            # SLA Realização
            if 'DENTRO_SLA' in row and pd.notna(row['DENTRO_SLA']):
                if row['DENTRO_SLA']:
                    sla_real_item = QTableWidgetItem('✅ Dentro')
                    sla_real_item.setForeground(QColor('#4CAF50'))
                else:
                    sla_real_item = QTableWidgetItem('❌ Fora')
                    sla_real_item.setForeground(QColor('#F44336'))
            else:
                sla_real_item = QTableWidgetItem('N/A')
            self.exams_table.setItem(i, 5, sla_real_item)

            # Data/Hora Entrega
            if 'DATA_ENTREGA_RESULTADO' in row and pd.notna(row['DATA_ENTREGA_RESULTADO']):
                entrega_item = QTableWidgetItem(row['DATA_ENTREGA_RESULTADO'].strftime('%d/%m/%Y %H:%M'))
            elif pd.notna(row['STATUS_PRELIMINAR']):
                entrega_item = QTableWidgetItem(row['STATUS_PRELIMINAR'].strftime('%d/%m/%Y %H:%M'))
            elif pd.notna(row['STATUS_APROVADO']):
                entrega_item = QTableWidgetItem(row['STATUS_APROVADO'].strftime('%d/%m/%Y %H:%M'))
            else:
                entrega_item = QTableWidgetItem('N/A')
            self.exams_table.setItem(i, 6, entrega_item)

            # Tempo Liberação (STATUS_ALAUDAR até entrega)
            if 'TEMPO_ENTREGA_MIN' in row and pd.notna(row['TEMPO_ENTREGA_MIN']):
                tempo_entrega = row['TEMPO_ENTREGA_MIN']
                if tempo_entrega > 1440:  # Mais de 24 horas
                    dias = tempo_entrega / 1440
                    tempo_text = f"{tempo_entrega:.1f} ({dias:.1f} dias)"
                else:
                    tempo_text = f"{tempo_entrega:.1f}"
                tempo_entrega_item = QTableWidgetItem(tempo_text)
            else:
                tempo_entrega_item = QTableWidgetItem('N/A')
            self.exams_table.setItem(i, 7, tempo_entrega_item)

            # SLA Entrega
            if 'DENTRO_SLA_ENTREGA' in row and pd.notna(row['DENTRO_SLA_ENTREGA']):
                if row['DENTRO_SLA_ENTREGA']:
                    sla_entrega_item = QTableWidgetItem('✅ Dentro')
                    sla_entrega_item.setForeground(QColor('#4CAF50'))
                else:
                    sla_entrega_item = QTableWidgetItem('❌ Fora')
                    sla_entrega_item.setForeground(QColor('#F44336'))
            else:
                sla_entrega_item = QTableWidgetItem('N/A')
            self.exams_table.setItem(i, 8, sla_entrega_item)

            # SLA Esperado
            if 'SLA_ESPERADO_DESC' in row and pd.notna(row['SLA_ESPERADO_DESC']):
                sla_desc_item = QTableWidgetItem(str(row['SLA_ESPERADO_DESC']))
            else:
                sla_desc_item = QTableWidgetItem('N/A')
            self.exams_table.setItem(i, 9, sla_desc_item)

            # Dias Úteis
            if 'DIAS_UTEIS_ENTREGA' in row and pd.notna(row['DIAS_UTEIS_ENTREGA']):
                dias_item = QTableWidgetItem(f"{int(row['DIAS_UTEIS_ENTREGA'])}")
            else:
                dias_item = QTableWidgetItem('N/A')
            self.exams_table.setItem(i, 10, dias_item)

    def export_to_excel(self):
        """Exporta histórico de exames para Excel"""
        try:
            search_term = self.search_input.text().strip()

            # Obter dados do paciente atual
            if search_term.isdigit():
                patient_data = self.df[self.df['SAME'].astype(str).str.contains(search_term, na=False)]
            else:
                patient_data = self.df[self.df['NOME_PACIENTE'].str.contains(search_term, case=False, na=False)]

            if len(patient_data) == 0:
                return

            # Preparar dados para exportação
            export_cols = [
                'SAME', 'NOME_PACIENTE', 'DATA_HORA_PRESCRICAO', 'GRUPO',
                'TIPO_ATENDIMENTO', 'STATUS_ALAUDAR', 'TEMPO_ATENDIMENTO_MIN',
                'DENTRO_SLA'
            ]

            # Adicionar colunas de entrega se existirem
            if 'DATA_ENTREGA_RESULTADO' in patient_data.columns:
                export_cols.append('DATA_ENTREGA_RESULTADO')
            if 'TEMPO_ENTREGA_MIN' in patient_data.columns:
                export_cols.append('TEMPO_ENTREGA_MIN')
            if 'DENTRO_SLA_ENTREGA' in patient_data.columns:
                export_cols.append('DENTRO_SLA_ENTREGA')
            if 'SLA_ESPERADO_DESC' in patient_data.columns:
                export_cols.append('SLA_ESPERADO_DESC')
            if 'DIAS_UTEIS_ENTREGA' in patient_data.columns:
                export_cols.append('DIAS_UTEIS_ENTREGA')

            # Filtrar apenas colunas existentes
            export_cols = [col for col in export_cols if col in patient_data.columns]
            export_df = patient_data[export_cols].copy()

            # Ordenar por data
            export_df = export_df.sort_values('DATA_HORA_PRESCRICAO', ascending=False)

            # Nome do arquivo
            same = patient_data.iloc[0]['SAME']
            nome = patient_data.iloc[0]['NOME_PACIENTE'].replace(' ', '_')[:30]
            filename = f"historico_exames_SAME_{int(same)}_{nome}.xlsx"

            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Salvar Histórico de Exames",
                build_runtime_file_path(filename),
                "Excel Files (*.xlsx)"
            )

            if file_path:
                export_df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Sucesso",
                                       f"Histórico exportado com sucesso!\n\n{file_path}")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar dados:\n{str(e)}")


class DashboardWindow(QMainWindow):
    """Janela separada com dashboards de SLA de Realização e SLA Entrega"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📊 Dashboard Analítico - CDI São José")
        self.setGeometry(100, 100, 1400, 900)

        # Widget central
        central_widget = QWidget()
        central_widget.setObjectName("dashboardRoot")
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Header
        header = QLabel("📊 Dashboard Analítico - Centro de Diagnóstico por Imagem")
        header.setStyleSheet("""
            font-size: 22px;
            font-weight: bold;
            color: #58a6ff;
            padding: 10px;
        """)
        main_layout.addWidget(header)

        subtitle = QLabel("Visão consolidada, com filtros rápidos e navegação por abas.")
        subtitle.setObjectName("subtitleLabel")
        main_layout.addWidget(subtitle)

        # Barra de busca de seções
        filter_bar = QWidget()
        filter_layout = QHBoxLayout(filter_bar)
        filter_layout.setContentsMargins(0, 0, 0, 0)
        filter_layout.setSpacing(8)

        filter_icon = QLabel("🔎")
        self.section_filter = QLineEdit()
        self.section_filter.setPlaceholderText("Filtrar seções do dashboard (ex: SLA, tabela, convênio)")
        self.section_filter.setClearButtonEnabled(True)
        self.section_filter.textChanged.connect(self.filter_sections)

        filter_layout.addWidget(filter_icon)
        filter_layout.addWidget(self.section_filter, stretch=1)
        main_layout.addWidget(filter_bar)

        # Tabs
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.currentChanged.connect(self.on_tab_changed)
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #30363d;
                background-color: #0d1117;
                border-radius: 6px;
            }
            QTabBar::tab {
                background-color: #1f2430;
                color: #c9d1d9;
                padding: 10px 20px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-weight: 500;
                font-size: 14px;
            }
            QTabBar::tab:selected {
                background-color: #2a9d8f;
                color: white;
            }
            QTabBar::tab:hover {
                background-color: #2b3340;
            }
        """)

        # Tab 1: SLA de Atendimento (Laudar)
        self.tab_laudar = QWidget()
        self.tab_laudar_layout = QVBoxLayout(self.tab_laudar)
        scroll_laudar = QScrollArea()
        scroll_laudar.setWidgetResizable(True)
        scroll_laudar.setFrameShape(QFrame.NoFrame)
        self.laudar_content = QWidget()
        self.laudar_content_layout = QVBoxLayout(self.laudar_content)
        scroll_laudar.setWidget(self.laudar_content)
        self.tab_laudar_layout.addWidget(scroll_laudar)

        # Tab 2: SLA de Entrega de Resultado
        self.tab_entrega = QWidget()
        self.tab_entrega_layout = QVBoxLayout(self.tab_entrega)
        scroll_entrega = QScrollArea()
        scroll_entrega.setWidgetResizable(True)
        scroll_entrega.setFrameShape(QFrame.NoFrame)
        self.entrega_content = QWidget()
        self.entrega_content_layout = QVBoxLayout(self.entrega_content)
        scroll_entrega.setWidget(self.entrega_content)
        self.tab_entrega_layout.addWidget(scroll_entrega)

        # Tab 3: Convênios
        self.tab_convenios = QWidget()
        self.tab_convenios_layout = QVBoxLayout(self.tab_convenios)
        scroll_convenios = QScrollArea()
        scroll_convenios.setWidgetResizable(True)
        scroll_convenios.setFrameShape(QFrame.NoFrame)
        self.convenios_content = QWidget()
        self.convenios_content_layout = QVBoxLayout(self.convenios_content)
        scroll_convenios.setWidget(self.convenios_content)
        self.tab_convenios_layout.addWidget(scroll_convenios)

        self.tabs.addTab(self.tab_laudar, "⚡ SLA de Atendimento (Laudar)")
        self.tabs.addTab(self.tab_entrega, "📦 SLA de Entrega de Resultado")
        self.tabs.addTab(self.tab_convenios, "🏥 Convênios")

        main_layout.addWidget(self.tabs)

        # Botões
        button_layout = QHBoxLayout()

        self.export_btn = QPushButton("💾 Exportar Dashboard")
        self.export_btn.setObjectName("primaryButton")
        self.export_btn.clicked.connect(self.export_dashboard)

        self.close_btn = QPushButton("✖ Fechar")
        self.close_btn.clicked.connect(self.close)

        button_layout.addWidget(self.export_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.close_btn)

        main_layout.addLayout(button_layout)

        # Aplicar tema escuro
        self.apply_dark_theme()

    def apply_dark_theme(self):
        """Aplica tema escuro"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #0b0f14;
            }
            QWidget {
                background-color: transparent;
                color: #c9d1d9;
                font-family: 'Manrope', 'Montserrat', 'Segoe UI', sans-serif;
            }
            QWidget#dashboardRoot {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #0b0f14, stop:0.5 #0f1621, stop:1 #0b121a);
                border-radius: 12px;
            }
            QLabel#subtitleLabel {
                color: #8b949e;
                font-size: 13px;
                padding: 0 10px 8px 10px;
            }
            QLineEdit {
                background-color: #111826;
                border: 1px solid #243041;
                border-radius: 8px;
                padding: 8px 12px;
                color: #e6edf3;
                font-size: 13px;
            }
            QLineEdit:focus {
                border-color: #2a9d8f;
            }
            QPushButton {
                background-color: #1f2430;
                border: 1px solid #2b3340;
                border-radius: 8px;
                padding: 8px 16px;
                color: #c9d1d9;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #2b3340;
                border-color: #3b4758;
            }
            QPushButton#primaryButton {
                background-color: #2a9d8f;
                border-color: #2a9d8f;
                color: #0b0f14;
            }
            QPushButton#primaryButton:hover {
                background-color: #34b3a3;
                border-color: #34b3a3;
            }
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QGroupBox {
                background-color: #111826;
                border: 1px solid #243041;
                border-radius: 12px;
                margin-top: 12px;
                padding-top: 16px;
                font-weight: 600;
                color: #9cc7ff;
            }
            QGroupBox:hover {
                border-color: #2a9d8f;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }
            QTableWidget {
                background-color: #0d1117;
                alternate-background-color: #111826;
                border: 1px solid #243041;
                border-radius: 8px;
                gridline-color: #243041;
            }
            QTableWidget::item {
                padding: 6px;
                color: #c9d1d9;
            }
            QTableWidget::item:selected {
                background-color: #2a9d8f;
                color: #0b0f14;
            }
            QHeaderView::section {
                background-color: #0f1621;
                color: #9cc7ff;
                padding: 8px;
                border: none;
                border-right: 1px solid #243041;
                border-bottom: 1px solid #243041;
                font-weight: 600;
            }
        """)

    def populate_laudar_dashboard(self, widgets_list):
        """Popula dashboard de SLA de Realização"""
        # Limpar conteúdo existente
        while self.laudar_content_layout.count():
            item = self.laudar_content_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # Adicionar widgets
        for widget in widgets_list:
            self.laudar_content_layout.addWidget(widget)
        self._apply_card_effects(self.tab_laudar)

    def populate_entrega_dashboard(self, widgets_list):
        """Popula dashboard de SLA Entrega"""
        # Limpar conteúdo existente
        while self.entrega_content_layout.count():
            item = self.entrega_content_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # Adicionar widgets
        for widget in widgets_list:
            self.entrega_content_layout.addWidget(widget)
        self._apply_card_effects(self.tab_entrega)

    def populate_convenio_dashboard(self, items_list):
        """Popula dashboard de Convênios"""
        while self.convenios_content_layout.count():
            item = self.convenios_content_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                self._clear_layout(item.layout())

        for item in items_list:
            if isinstance(item, QLayout):
                self.convenios_content_layout.addLayout(item)
            else:
                self.convenios_content_layout.addWidget(item)
        self._apply_card_effects(self.tab_convenios)

    def _clear_layout(self, layout):
        """Limpa layouts aninhados"""
        while layout.count():
            item = layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                self._clear_layout(item.layout())

    def _apply_card_effects(self, root_widget):
        """Aplica sombra e estilo de card aos grupos"""
        for group in root_widget.findChildren(QGroupBox):
            group.setProperty("card", True)
            if group.graphicsEffect() is None:
                shadow = QGraphicsDropShadowEffect(group)
                shadow.setBlurRadius(18)
                shadow.setOffset(0, 6)
                shadow.setColor(QColor(0, 0, 0, 140))
                group.setGraphicsEffect(shadow)

    def filter_sections(self, text):
        """Filtra seções por título na aba atual"""
        query = text.strip().lower()
        current_tab = self.tabs.currentWidget()
        if current_tab is None:
            return
        for group in current_tab.findChildren(QGroupBox):
            title = group.title().lower()
            group.setVisible(not query or query in title)

    def on_tab_changed(self):
        """Atualiza filtro e animação ao trocar de aba"""
        if hasattr(self, 'section_filter'):
            self.filter_sections(self.section_filter.text())

    def export_dashboard(self):
        """Exporta dados do dashboard"""
        QMessageBox.information(self, "Exportar", "Funcionalidade de exportação em desenvolvimento.")


class AIAnalysisWindow(QMainWindow):
    """Janela separada para exibir análise de IA"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🤖 Análise Estratégica com IA - CDI São José")
        self.setGeometry(150, 150, 1000, 700)

        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)

        # Header
        header = QLabel("🤖 Análise Estratégica Gerada por IA")
        header.setStyleSheet("""
            font-size: 20px;
            font-weight: bold;
            color: #58a6ff;
            padding: 10px;
        """)
        layout.addWidget(header)

        # Área de texto
        self.text_area = QTextEdit()
        self.text_area.setReadOnly(True)
        self.text_area.setStyleSheet("""
            QTextEdit {
                background-color: #0d1117;
                border: 1px solid #30363d;
                border-radius: 6px;
                padding: 20px;
                color: #c9d1d9;
                font-size: 14px;
                line-height: 1.7;
            }
        """)
        layout.addWidget(self.text_area)

        # Botões
        button_layout = QHBoxLayout()

        self.export_btn = QPushButton("💾 Exportar Análise")
        self.export_btn.setObjectName("primaryButton")
        self.export_btn.clicked.connect(self.export_analysis)

        self.copy_btn = QPushButton("📋 Copiar para Clipboard")
        self.copy_btn.clicked.connect(self.copy_to_clipboard)

        self.close_btn = QPushButton("✖ Fechar")
        self.close_btn.clicked.connect(self.close)

        button_layout.addWidget(self.export_btn)
        button_layout.addWidget(self.copy_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.close_btn)

        layout.addLayout(button_layout)

        # Aplicar tema escuro
        self.setStyleSheet("""
            QMainWindow {
                background-color: #0d1117;
            }
            QWidget {
                background-color: #0d1117;
                color: #c9d1d9;
            }
            QPushButton {
                background-color: #21262d;
                border: 1px solid #30363d;
                border-radius: 6px;
                padding: 8px 16px;
                color: #c9d1d9;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #30363d;
                border-color: #58a6ff;
            }
            QPushButton#primaryButton {
                background-color: #238636;
                border-color: #238636;
            }
            QPushButton#primaryButton:hover {
                background-color: #2ea043;
            }
        """)

    def set_analysis(self, text):
        """Define o texto da análise"""
        self.text_area.setMarkdown(text)

    def set_loading(self):
        """Mostra estado de carregamento"""
        self.text_area.setText("🔄 Gerando análise estratégica com IA...\n\nAguarde, isso pode levar alguns segundos.")
        self.export_btn.setEnabled(False)
        self.copy_btn.setEnabled(False)

    def set_error(self, error_msg):
        """Mostra mensagem de erro"""
        self.text_area.setText(f"❌ Erro ao gerar análise:\n\n{error_msg}")
        self.export_btn.setEnabled(False)
        self.copy_btn.setEnabled(False)

    def set_complete(self, text):
        """Análise completa"""
        self.text_area.setMarkdown(text)
        self.export_btn.setEnabled(True)
        self.copy_btn.setEnabled(True)

    def export_analysis(self):
        """Exporta análise para arquivo"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar Análise Estratégica",
            build_runtime_file_path(
                f"analise_estrategica_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            ),
            "Text Files (*.txt);;Markdown Files (*.md)"
        )

        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.text_area.toPlainText())
                QMessageBox.information(self, "Sucesso", f"Análise exportada:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao exportar:\n{str(e)}")

    def copy_to_clipboard(self):
        """Copia análise para clipboard"""
        clipboard = QApplication.clipboard()
        clipboard.setText(self.text_area.toPlainText())
        QMessageBox.information(self, "Sucesso", "Análise copiada para o clipboard!")


class AIAnalysisThread(QThread):
    """Thread para gerar análise estratégica com IA"""
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, stats_gerais, stats_entrega, analise_grupo, analise_tipo, api_type, api_key, api_url):
        super().__init__()
        self.stats_gerais = stats_gerais
        self.stats_entrega = stats_entrega
        self.analise_grupo = analise_grupo
        self.analise_tipo = analise_tipo
        self.api_type = api_type
        self.api_key = api_key
        self.api_url = api_url

    def run(self):
        try:
            # Preparar dados para análise
            prompt = self.criar_prompt_analise()

            # Gerar análise com IA
            if self.api_type == "OpenAI":
                analise = self.gerar_analise_openai(prompt)
            elif self.api_type == "LM Studio":
                analise = self.gerar_analise_lmstudio(prompt)
            else:
                analise = "API não configurada corretamente."

            self.finished.emit(analise)

        except Exception as e:
            self.error.emit(f"Erro ao gerar análise: {str(e)}")

    def criar_prompt_analise(self):
        """Cria prompt estruturado para análise"""
        # Análise por grupo
        grupo_stats = []
        for grupo, row in self.analise_grupo.iterrows():
            tempo_medio = row[('TEMPO_ATENDIMENTO_MIN', 'mean')]
            total = int(row[('TEMPO_ATENDIMENTO_MIN', 'count')])
            dentro_sla = int(row[('DENTRO_SLA', 'sum')])
            perc_sla = (dentro_sla / total * 100) if total > 0 else 0
            grupo_stats.append(f"- {grupo}: {total} exames, tempo médio {tempo_medio:.1f}min, {perc_sla:.1f}% no SLA")

        # Análise por tipo
        tipo_stats = []
        for tipo, row in self.analise_tipo.iterrows():
            tempo_medio = row[('TEMPO_ATENDIMENTO_MIN', 'mean')]
            total = int(row[('TEMPO_ATENDIMENTO_MIN', 'count')])
            tipo_stats.append(f"- {tipo}: {total} exames, tempo médio {tempo_medio:.1f}min")

        prompt = f"""Você é um consultor especializado em gestão hospitalar e análise de indicadores de qualidade em serviços de diagnóstico por imagem.

Analise os dados abaixo do Centro de Diagnóstico por Imagem da Casa de Saúde São José e gere uma análise estratégica executiva.

## DADOS GERAIS - SLA DE ATENDIMENTO (LAUDAR)
- Total de exames: {self.stats_gerais['total_exames']:,}
- Tempo médio de laudar: {self.stats_gerais['tempo_medio']:.1f} minutos
- Meta SLA: ≤60 minutos
- Taxa de cumprimento: {self.stats_gerais['percentual_sla']:.1f}%
- Exames fora do prazo: {self.stats_gerais['fora_sla']:,}

## DADOS POR MODALIDADE (ATENDIMENTO)
{chr(10).join(grupo_stats)}

## DADOS POR TIPO DE ATENDIMENTO
{chr(10).join(tipo_stats)}

## SLA DE ENTREGA DE RESULTADO
- Total de exames com resultado: {self.stats_entrega['total_exames']:,}
- Tempo médio de entrega: {self.stats_entrega['tempo_medio_entrega']:.1f} horas
- Taxa de cumprimento SLA: {self.stats_entrega['percentual_sla_entrega']:.1f}%
- Exames fora do prazo de entrega: {self.stats_entrega['fora_sla_entrega']:,}

## ANÁLISE SOLICITADA

Forneça uma análise estratégica estruturada com:

1. **Resumo Executivo** (2-3 parágrafos)
2. **Principais Problemas Identificados** (liste os 3-5 principais gargalos)
3. **Análise por Modalidade** (quais modalidades têm pior desempenho e por quê)
4. **Análise por Tipo de Atendimento** (diferenças entre PA, Internado, Externo)
5. **Recomendações Prioritárias** (5-7 ações concretas com impacto esperado)
6. **Indicadores de Acompanhamento** (KPIs que devem ser monitorados)

Use linguagem técnica mas acessível para gestores hospitalares. Seja direto e objetivo."""

        return prompt

    def gerar_analise_openai(self, prompt):
        """Gera análise usando OpenAI API"""
        if not OPENAI_AVAILABLE:
            return "Biblioteca OpenAI não instalada. Execute: pip install openai"

        try:
            client = openai.OpenAI(api_key=self.api_key)
            response = client.chat.completions.create(
                model="gpt-4o",  # ou "gpt-3.5-turbo" para economia
                messages=[
                    {"role": "system", "content": "Você é um consultor especializado em gestão hospitalar e análise de indicadores de qualidade."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=2000
            )
            return response.choices[0].message.content

        except Exception as e:
            return f"Erro ao conectar com OpenAI: {str(e)}"

    def gerar_analise_lmstudio(self, prompt):
        """Gera análise usando LM Studio local"""
        if not REQUESTS_AVAILABLE:
            return "Biblioteca requests não instalada. Execute: pip install requests"

        try:
            # LM Studio endpoint é /v1/chat/completions, mas a URL base já deve incluir /v1
            # Então se api_url = "http://localhost:1234/v1", usar apenas api_url
            base_url = self.api_url.rstrip('/')

            # Se não termina com /v1, adicionar
            if not base_url.endswith('/v1'):
                endpoint = f"{base_url}/v1/chat/completions"
            else:
                endpoint = f"{base_url}/chat/completions"

            response = requests.post(
                endpoint,
                json={
                    "model": "local-model",  # LM Studio ignora este campo
                    "messages": [
                        {"role": "system", "content": "Você é um consultor especializado em gestão hospitalar e análise de indicadores de qualidade."},
                        {"role": "user", "content": prompt}
                    ],
                    "temperature": 0.7,
                    "max_tokens": 2500,
                    "stream": False
                },
                timeout=180  # Aumentado para 3 minutos
            )

            if response.status_code == 200:
                result = response.json()
                return result['choices'][0]['message']['content']
            else:
                return f"Erro na requisição LM Studio: {response.status_code}\n\nResposta: {response.text}"

        except requests.exceptions.Timeout:
            return f"Timeout ao conectar com LM Studio.\n\nVerifique se:\n1. LM Studio está rodando em {self.api_url}\n2. Um modelo está carregado\n3. O servidor está configurado para aceitar requisições\n4. O modelo não está sobrecarregado"
        except requests.exceptions.ConnectionError:
            return f"Erro de conexão com LM Studio em {self.api_url}\n\nVerifique se:\n1. LM Studio está rodando\n2. O servidor local está ativo\n3. A porta está correta (normalmente 1234)"
        except Exception as e:
            return f"Erro ao conectar com LM Studio: {str(e)}\n\nURL usada: {self.api_url}"


class DataProcessor(QThread):
    """Thread para processar dados sem bloquear a UI"""
    finished = Signal(dict)
    error = Signal(str)

    def __init__(self, file_path, start_date, end_date):
        super().__init__()
        self.file_path = file_path
        self.start_date = start_date
        self.end_date = end_date

    def calcular_sla_entrega(self, row):
        """Calcula o SLA de entrega em minutos baseado em GRUPO e TIPO_ATENDIMENTO"""
        grupo = str(row['GRUPO']).upper() if pd.notna(row['GRUPO']) else ''
        tipo = str(row['TIPO_ATENDIMENTO']).upper() if pd.notna(row['TIPO_ATENDIMENTO']) else ''

        # Tomografia
        if 'TOMOGRAFIA' in grupo:
            if 'PRONTO ATENDIMENTO' in tipo or 'PRONTO' in tipo:
                return 72  # 1.2 horas (1h12min)
            elif 'INTERNADO' in tipo:
                return 1440  # 24 horas (24 * 60)
            elif 'EXTERNO' in tipo:
                return 999999  # Marcador para 5 dias úteis
            else:
                return 1440  # Default: 24 horas

        # Ressonância
        if 'RESSONÂNCIA' in grupo or 'RESSONANCIA' in grupo or 'MAGNÉTICA' in grupo:
            if 'PRONTO ATENDIMENTO' in tipo or 'PRONTO' in tipo:
                return 72  # 1.2 horas (1h12min)
            elif 'INTERNADO' in tipo:
                return 1440  # 24 horas
            elif 'EXTERNO' in tipo:
                return 999999  # Marcador para 5 dias úteis
            else:
                return 1440  # Default: 24 horas

        # Raio X - sempre 3 dias úteis
        if 'RAIO' in grupo:
            return 999998  # Marcador para 3 dias úteis

        # Mamografia - 5 dias úteis
        if 'MAMOGRAFIA' in grupo:
            return 999999  # Marcador para 5 dias úteis

        # Medicina Nuclear - 5 dias úteis
        if 'MEDICINA NUCLEAR' in grupo or 'NUCLEAR' in grupo:
            return 999999  # Marcador para 5 dias úteis

        # Ultrassonografia - EXTERNO: 3 dias úteis
        if 'ULTRASSOM' in grupo or 'ULTRASOM' in grupo or 'ECOGRAFIA' in grupo:
            if 'EXTERNO' in tipo:
                return 999998  # Marcador para 3 dias úteis
            else:
                return 1440  # Default: 24 horas para outros tipos

        # Default: 24 horas
        return 1440

    def calcular_dias_uteis(self, row):
        """Calcula dias úteis entre laudo (STATUS_ALAUDAR) e entrega (excluindo sábados e domingos)"""
        if pd.isna(row['STATUS_ALAUDAR']) or pd.isna(row['DATA_ENTREGA_RESULTADO']):
            return None

        data_inicio = row['STATUS_ALAUDAR'].date()
        data_fim = row['DATA_ENTREGA_RESULTADO'].date()

        # D0 é o dia do laudo
        dias_uteis = 0
        data_atual = data_inicio

        while data_atual <= data_fim:
            # 0 = segunda, 6 = domingo
            if data_atual.weekday() < 5:  # Segunda a sexta
                dias_uteis += 1
            data_atual += timedelta(days=1)

        return dias_uteis - 1 if dias_uteis > 0 else 0  # Descontar D0

    def run(self):
        try:
            # Carregar dados COMPLETOS da planilha (sem filtro de data ainda)
            df_completo = pd.read_excel(self.file_path)

            # Converter datas com formato brasileiro (dia-mês-ano)
            df_completo['DATA_HORA_PRESCRICAO'] = pd.to_datetime(df_completo['DATA_HORA_PRESCRICAO'],
                                                         format='%d-%m-%Y %H:%M',
                                                         errors='coerce')
            df_completo['STATUS_ALAUDAR'] = pd.to_datetime(df_completo['STATUS_ALAUDAR'],
                                                   format='%d-%m-%Y %H:%M',
                                                   errors='coerce')
            df_completo['STATUS_PRELIMINAR'] = pd.to_datetime(df_completo['STATUS_PRELIMINAR'],
                                                      format='%d-%m-%Y %H:%M',
                                                      errors='coerce')
            df_completo['STATUS_APROVADO'] = pd.to_datetime(df_completo['STATUS_APROVADO'],
                                                    format='%d-%m-%Y %H:%M',
                                                    errors='coerce')

            # Filtrar por período APENAS para análises principais (SLA de Realização e SLA Entrega)
            mask = (df_completo['DATA_HORA_PRESCRICAO'].dt.date >= self.start_date) & \
                   (df_completo['DATA_HORA_PRESCRICAO'].dt.date <= self.end_date)
            df = df_completo[mask]

            # Filtrar GRUPO OUTROS, ECOCARDIOGRAMA, DENSITOMETRIA e TIPO_ATENDIMENTO AUDITORIA
            df = df[~df['GRUPO'].str.upper().str.contains('OUTROS', na=False)]
            df = df[~df['GRUPO'].str.upper().str.contains('ECOCARDIOGRAMA', na=False)]
            df = df[~df['GRUPO'].str.upper().str.contains('DENSITOMETRIA', na=False)]
            df = df[~df['TIPO_ATENDIMENTO'].str.upper().str.contains('AUDITORIA', na=False)]

            # Calcular tempo de atendimento em minutos (para laudo)
            df['TEMPO_ATENDIMENTO_MIN'] = (df['STATUS_ALAUDAR'] - df['DATA_HORA_PRESCRICAO']).dt.total_seconds() / 60

            # Filtrar apenas registros válidos
            df = df[df['TEMPO_ATENDIMENTO_MIN'].notna()]
            df = df[df['TEMPO_ATENDIMENTO_MIN'] >= 0]

            # Classificar se atende SLA de atendimento (60 minutos) - ANTES da análise longitudinal
            df['DENTRO_SLA'] = df['TEMPO_ATENDIMENTO_MIN'] <= 60

            # ========== ANÁLISE LONGITUDINAL DIÁRIA - SLA LAUDAR ==========
            # Criar coluna de DIA_PLANTAO que considera plantão de 7h às 7h
            def calcular_dia_plantao(data_hora):
                """Retorna a data do plantão (7h-7h do dia seguinte)"""
                if pd.isna(data_hora):
                    return None
                # Se hora < 7h da manhã, pertence ao plantão do dia anterior
                if data_hora.hour < 7:
                    return (data_hora - timedelta(days=1)).date()
                else:
                    return data_hora.date()

            df['DIA_PLANTAO'] = df['DATA_HORA_PRESCRICAO'].apply(calcular_dia_plantao)

            # Mapear modalidades para análise longitudinal
            def classificar_modalidade_laudar(grupo):
                grupo_upper = str(grupo).upper()
                if 'TOMOGRAFIA' in grupo_upper:
                    return 'TC'
                elif 'RESSONÂNCIA' in grupo_upper or 'RESSONANCIA' in grupo_upper or 'MAGNÉTICA' in grupo_upper:
                    return 'RM'
                elif 'ULTRASSOM' in grupo_upper or 'ULTRASOM' in grupo_upper or 'ECOGRAFIA' in grupo_upper:
                    return 'US'
                elif 'RAIO' in grupo_upper:
                    return 'RX'
                elif 'MEDICINA NUCLEAR' in grupo_upper or 'NUCLEAR' in grupo_upper or 'CINTILOGRAFIA' in grupo_upper:
                    return 'MN'
                else:
                    return 'OUTROS'

            df['MODALIDADE'] = df['GRUPO'].apply(classificar_modalidade_laudar)

            # Agrupar por DIA_PLANTAO, MODALIDADE e TIPO_ATENDIMENTO
            analise_longitudinal_laudar = df.groupby(
                ['DIA_PLANTAO', 'MODALIDADE', 'TIPO_ATENDIMENTO']
            ).agg({
                'TEMPO_ATENDIMENTO_MIN': ['count', 'mean', 'median'],
                'DENTRO_SLA': 'sum'
            }).reset_index()

            # Renomear colunas
            analise_longitudinal_laudar.columns = [
                'DIA_PLANTAO', 'MODALIDADE', 'TIPO_ATENDIMENTO',
                'QTD_EXAMES', 'TEMPO_MEDIO', 'TEMPO_MEDIANO', 'DENTRO_SLA'
            ]

            # Calcular percentual de SLA
            analise_longitudinal_laudar['PERCENTUAL_SLA'] = (
                analise_longitudinal_laudar['DENTRO_SLA'] /
                analise_longitudinal_laudar['QTD_EXAMES'] * 100
            )

            # Formatar data para exibição
            analise_longitudinal_laudar['DIA_STR'] = pd.to_datetime(
                analise_longitudinal_laudar['DIA_PLANTAO']
            ).dt.strftime('%d/%m')

            # Ordenar por data
            analise_longitudinal_laudar = analise_longitudinal_laudar.sort_values('DIA_PLANTAO')

            # Calcular tempo de ENTREGA DE RESULTADO
            # Data inicial: STATUS_ALAUDAR (quando foi laudado)
            # Data final: STATUS_PRELIMINAR se existir, senão STATUS_APROVADO
            df['DATA_ENTREGA_RESULTADO'] = df['STATUS_PRELIMINAR'].fillna(df['STATUS_APROVADO'])
            df['TEMPO_ENTREGA_MIN'] = (df['DATA_ENTREGA_RESULTADO'] - df['STATUS_ALAUDAR']).dt.total_seconds() / 60

            # Calcular SLA de entrega baseado em GRUPO e TIPO_ATENDIMENTO
            df['SLA_ENTREGA_MIN'] = df.apply(self.calcular_sla_entrega, axis=1)
            df['DIAS_UTEIS_ENTREGA'] = df.apply(self.calcular_dias_uteis, axis=1)

            # Ajustar SLA para casos de dias úteis
            def ajustar_sla_dias_uteis(row):
                sla = row['SLA_ENTREGA_MIN']
                dias_uteis = row['DIAS_UTEIS_ENTREGA']

                # Se marcador 999999 = 5 dias úteis
                if sla == 999999:
                    return dias_uteis <= 5
                # Se marcador 999998 = 3 dias úteis
                elif sla == 999998:
                    return dias_uteis <= 3
                # Senão, comparar tempo em minutos
                else:
                    return row['TEMPO_ENTREGA_MIN'] <= sla

            df['DENTRO_SLA_ENTREGA'] = df.apply(ajustar_sla_dias_uteis, axis=1)

            # Criar coluna com descrição do SLA esperado
            def formatar_sla_esperado(row):
                sla = row['SLA_ENTREGA_MIN']
                if sla == 999999:
                    return "5 dias úteis"
                elif sla == 999998:
                    return "3 dias úteis"
                elif sla == 60:
                    return "1 hora"
                elif sla == 72:
                    return "1,2 horas"
                elif sla == 1440:
                    return "24 horas"
                else:
                    horas = sla / 60
                    if horas == int(horas):
                        return f"{int(horas)} horas"
                    else:
                        return f"{horas:.1f} horas"

            df['SLA_ESPERADO_DESC'] = df.apply(formatar_sla_esperado, axis=1)

            # Filtrar registros com dados de entrega válidos
            df_entrega = df[df['TEMPO_ENTREGA_MIN'].notna()].copy()

            # Análise por GRUPO (modalidade) - ATENDIMENTO
            analise_grupo = df.groupby('GRUPO').agg({
                'TEMPO_ATENDIMENTO_MIN': ['mean', 'median', 'min', 'max', 'count'],
                'DENTRO_SLA': 'sum'
            }).round(2)

            # Análise por TIPO_ATENDIMENTO - ATENDIMENTO
            analise_tipo = df.groupby('TIPO_ATENDIMENTO').agg({
                'TEMPO_ATENDIMENTO_MIN': ['mean', 'median', 'min', 'max', 'count'],
                'DENTRO_SLA': 'sum'
            }).round(2)

            # Análise combinada - ATENDIMENTO
            analise_combinada = df.groupby(['GRUPO', 'TIPO_ATENDIMENTO']).agg({
                'TEMPO_ATENDIMENTO_MIN': ['mean', 'count'],
                'DENTRO_SLA': 'sum'
            }).round(2)

            # Análise por GRUPO - ENTREGA
            analise_grupo_entrega = df_entrega.groupby('GRUPO').agg({
                'TEMPO_ENTREGA_MIN': ['mean', 'median', 'min', 'max', 'count'],
                'DENTRO_SLA_ENTREGA': 'sum',
                'DIAS_UTEIS_ENTREGA': 'mean'
            }).round(2)

            # Análise combinada - ENTREGA
            analise_combinada_entrega = df_entrega.groupby(['GRUPO', 'TIPO_ATENDIMENTO']).agg({
                'TEMPO_ENTREGA_MIN': ['mean', 'count'],
                'DENTRO_SLA_ENTREGA': 'sum',
                'DIAS_UTEIS_ENTREGA': 'mean'
            }).round(2)

            # Estatísticas gerais - ATENDIMENTO
            stats_gerais = {
                'total_exames': len(df),
                'tempo_medio': df['TEMPO_ATENDIMENTO_MIN'].mean(),
                'tempo_mediano': df['TEMPO_ATENDIMENTO_MIN'].median(),
                'tempo_min': df['TEMPO_ATENDIMENTO_MIN'].min(),
                'tempo_max': df['TEMPO_ATENDIMENTO_MIN'].max(),
                'dentro_sla': df['DENTRO_SLA'].sum(),
                'fora_sla': (~df['DENTRO_SLA']).sum(),
                'percentual_sla': (df['DENTRO_SLA'].sum() / len(df) * 100) if len(df) > 0 else 0
            }

            # Estatísticas gerais - ENTREGA
            stats_entrega = {
                'total_exames': len(df_entrega),
                'tempo_medio_entrega': df_entrega['TEMPO_ENTREGA_MIN'].mean() / 60,  # em horas
                'dias_uteis_medio': df_entrega['DIAS_UTEIS_ENTREGA'].mean(),
                'dentro_sla_entrega': df_entrega['DENTRO_SLA_ENTREGA'].sum(),
                'fora_sla_entrega': (~df_entrega['DENTRO_SLA_ENTREGA']).sum(),
                'percentual_sla_entrega': (df_entrega['DENTRO_SLA_ENTREGA'].sum() / len(df_entrega) * 100) if len(df_entrega) > 0 else 0
            }

            # Distribuição por faixa de tempo
            bins = [0, 15, 30, 45, 60, 90, 120, float('inf')]
            labels = ['0-15min', '15-30min', '30-45min', '45-60min', '60-90min', '90-120min', '>120min']
            df['FAIXA_TEMPO'] = pd.cut(df['TEMPO_ATENDIMENTO_MIN'], bins=bins, labels=labels)
            distribuicao_tempo = df['FAIXA_TEMPO'].value_counts().sort_index()

            # ========== ANÁLISE LONGITUDINAL MENSAL - SLA ENTREGA ==========
            # Para análise longitudinal, usar dados COMPLETOS da planilha (df_completo)
            # independente do filtro de data da interface

            # Criar cópia dos dados completos para análise longitudinal
            df_long_completo = df_completo.copy()

            # Aplicar mesmos filtros de GRUPO e TIPO_ATENDIMENTO
            df_long_completo = df_long_completo[~df_long_completo['GRUPO'].str.upper().str.contains('OUTROS', na=False)]
            df_long_completo = df_long_completo[~df_long_completo['GRUPO'].str.upper().str.contains('ECOCARDIOGRAMA', na=False)]
            df_long_completo = df_long_completo[~df_long_completo['GRUPO'].str.upper().str.contains('DENSITOMETRIA', na=False)]
            df_long_completo = df_long_completo[~df_long_completo['TIPO_ATENDIMENTO'].str.upper().str.contains('AUDITORIA', na=False)]

            # Preparar dados de entrega para análise longitudinal
            df_long_completo['DATA_ENTREGA_RESULTADO'] = df_long_completo['STATUS_PRELIMINAR'].fillna(df_long_completo['STATUS_APROVADO'])
            df_long_completo['TEMPO_ENTREGA_MIN'] = (df_long_completo['DATA_ENTREGA_RESULTADO'] - df_long_completo['STATUS_ALAUDAR']).dt.total_seconds() / 60

            # Calcular SLA de entrega
            df_long_completo['SLA_ENTREGA_MIN'] = df_long_completo.apply(self.calcular_sla_entrega, axis=1)

            # Calcular dias úteis quando necessário
            df_long_completo['DIAS_UTEIS'] = df_long_completo.apply(self.calcular_dias_uteis, axis=1)

            # Determinar se está dentro do SLA
            def verificar_sla_entrega(row):
                if pd.isna(row['TEMPO_ENTREGA_MIN']):
                    return None
                if row['SLA_ENTREGA_MIN'] == 999999:  # 5 dias úteis
                    return row['DIAS_UTEIS'] <= 5 if pd.notna(row['DIAS_UTEIS']) else None
                elif row['SLA_ENTREGA_MIN'] == 999998:  # 3 dias úteis
                    return row['DIAS_UTEIS'] <= 3 if pd.notna(row['DIAS_UTEIS']) else None
                else:
                    return row['TEMPO_ENTREGA_MIN'] <= row['SLA_ENTREGA_MIN']

            df_long_completo['DENTRO_SLA_ENTREGA'] = df_long_completo.apply(verificar_sla_entrega, axis=1)

            # Filtrar apenas registros com dados válidos
            df_entrega_long = df_long_completo[
                df_long_completo['TEMPO_ENTREGA_MIN'].notna() &
                df_long_completo['DENTRO_SLA_ENTREGA'].notna()
            ].copy()

            # Extrair mês/ano da data de prescrição
            df_entrega_long['ANO_MES'] = df_entrega_long['DATA_HORA_PRESCRICAO'].dt.to_period('M')

            # Calcular o mês selecionado pelo usuário
            mes_selecionado = pd.Period(self.end_date, freq='M')

            # Calcular os 3 meses ANTERIORES + o mês atual = 4 meses no total
            mes_3_antes = mes_selecionado - 3

            # Filtrar os 3 meses anteriores + o mês vigente (total de 4 meses)
            df_entrega_copy = df_entrega_long[
                (df_entrega_long['ANO_MES'] >= mes_3_antes) &
                (df_entrega_long['ANO_MES'] <= mes_selecionado)
            ].copy()

            # Mapear modalidades
            def classificar_modalidade(grupo):
                grupo_upper = str(grupo).upper()
                if 'TOMOGRAFIA' in grupo_upper:
                    return 'TC'
                elif 'RESSONÂNCIA' in grupo_upper or 'RESSONANCIA' in grupo_upper or 'MAGNÉTICA' in grupo_upper:
                    return 'RM'
                elif 'ULTRASSOM' in grupo_upper or 'ULTRASOM' in grupo_upper or 'ECOGRAFIA' in grupo_upper:
                    return 'US'
                elif 'RAIO' in grupo_upper:
                    return 'RX'
                elif 'MEDICINA NUCLEAR' in grupo_upper or 'NUCLEAR' in grupo_upper or 'CINTILOGRAFIA' in grupo_upper:
                    return 'MN'
                else:
                    return 'OUTROS'

            df_entrega_copy['MODALIDADE'] = df_entrega_copy['GRUPO'].apply(classificar_modalidade)

            # Agrupar por mês, modalidade e tipo de atendimento
            analise_longitudinal_entrega = df_entrega_copy.groupby(
                ['ANO_MES', 'MODALIDADE', 'TIPO_ATENDIMENTO']
            ).agg({
                'TEMPO_ENTREGA_MIN': ['count', 'mean'],
                'DENTRO_SLA_ENTREGA': 'sum'
            }).reset_index()

            # Ordenar por ano/mês cronologicamente
            analise_longitudinal_entrega = analise_longitudinal_entrega.sort_values('ANO_MES')

            # Calcular percentual de SLA
            analise_longitudinal_entrega['PERCENTUAL_SLA'] = (
                analise_longitudinal_entrega[('DENTRO_SLA_ENTREGA', 'sum')] /
                analise_longitudinal_entrega[('TEMPO_ENTREGA_MIN', 'count')] * 100
            )

            # Renomear colunas para facilitar acesso
            analise_longitudinal_entrega.columns = [
                'ANO_MES', 'MODALIDADE', 'TIPO_ATENDIMENTO',
                'QTD_EXAMES', 'TEMPO_MEDIO', 'DENTRO_SLA', 'PERCENTUAL_SLA'
            ]

            # Converter period para string formatado (em português)
            meses_pt = {
                'Jan': 'Jan', 'Feb': 'Fev', 'Mar': 'Mar', 'Apr': 'Abr',
                'May': 'Mai', 'Jun': 'Jun', 'Jul': 'Jul', 'Aug': 'Ago',
                'Sep': 'Set', 'Oct': 'Out', 'Nov': 'Nov', 'Dec': 'Dez'
            }

            def formatar_mes_ano(periodo):
                if pd.isna(periodo):
                    return ''
                mes_ano = periodo.strftime('%b/%Y')
                # Substituir nome do mês em inglês por português
                for en, pt in meses_pt.items():
                    if mes_ano.startswith(en):
                        return mes_ano.replace(en, pt)
                return mes_ano

            analise_longitudinal_entrega['ANO_MES_STR'] = analise_longitudinal_entrega['ANO_MES'].apply(formatar_mes_ano)

            resultado = {
                'df': df,
                'df_entrega': df_entrega,
                'analise_grupo': analise_grupo,
                'analise_tipo': analise_tipo,
                'analise_combinada': analise_combinada,
                'analise_grupo_entrega': analise_grupo_entrega,
                'analise_combinada_entrega': analise_combinada_entrega,
                'stats_gerais': stats_gerais,
                'stats_entrega': stats_entrega,
                'distribuicao_tempo': distribuicao_tempo,
                'analise_longitudinal_entrega': analise_longitudinal_entrega,
                'analise_longitudinal_laudar': analise_longitudinal_laudar
            }

            self.finished.emit(resultado)

        except Exception as e:
            self.error.emit(str(e))


class StatCard(QFrame):
    """Card para exibir estatística"""
    def __init__(self, title, value, subtitle="", color="#2196F3"):
        super().__init__()
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        self.setStyleSheet(f"""
            StatCard {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {color}, stop:1 {self.darken_color(color)});
                border-radius: 10px;
                padding: 15px;
            }}
        """)

        layout = QVBoxLayout()

        title_label = QLabel(title)
        title_label.setStyleSheet("color: white; font-size: 12px; font-weight: normal;")

        value_label = QLabel(str(value))
        value_label.setStyleSheet("color: white; font-size: 32px; font-weight: bold;")

        subtitle_label = QLabel(subtitle)
        subtitle_label.setStyleSheet("color: rgba(255,255,255,0.8); font-size: 11px;")

        layout.addWidget(title_label)
        layout.addWidget(value_label)
        if subtitle:
            layout.addWidget(subtitle_label)
        layout.addStretch()

        self.setLayout(layout)
        self.setMinimumHeight(120)

    def darken_color(self, hex_color):
        """Escurece uma cor hex"""
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        r, g, b = int(r * 0.8), int(g * 0.8), int(b * 0.8)
        return f"#{r:02x}{g:02x}{b:02x}"


class MatplotlibChart(QWidget):
    """Widget para gráficos matplotlib"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.figure = Figure(facecolor='#1e1e1e', figsize=(8, 6))
        self.canvas = FigureCanvasQTAgg(self.figure)
        layout = QVBoxLayout()
        layout.addWidget(self.canvas)
        layout.setContentsMargins(10, 10, 10, 10)
        self.setLayout(layout)
        self.setMinimumHeight(400)

    def clear(self):
        self.figure.clear()
        self.canvas.draw()


class UltrasoundAnalysisWindow(QMainWindow):
    """Janela de Análise Estratégica de Ultrassonografia"""
    def __init__(self, df, start_date, end_date, unidade_nome, parent=None):
        super().__init__(parent)
        self.unidade_nome = unidade_nome
        self.setWindowTitle(f"🔬 Análise Estratégica de Ultrassonografia - {unidade_nome}")
        self.setGeometry(50, 50, 1700, 1000)

        # Armazenar período
        self.start_date = start_date
        self.end_date = end_date
        self.capacidade_diaria = 80  # Capacidade padrão de exames/dia

        # Filtrar apenas ULTRASSOM do dataframe recebido
        self.df_us = self.filter_ultrasound_data(df)

        self.setup_ui()
        self.apply_dark_theme()

        # Gerar análises automaticamente se houver dados
        if self.df_us is not None and len(self.df_us) > 0:
            self.generate_all_analyses()

    def filter_ultrasound_data(self, df):
        """Filtra apenas exames de ultrassonografia"""
        if df is None or len(df) == 0:
            return None

        # Filtrar apenas GRUPO que contenha ULTRASSOM (case insensitive)
        mask_us = df['GRUPO'].str.upper().str.contains('ULTRASSOM', na=False)
        df_us = df[mask_us].copy()

        if len(df_us) == 0:
            return None

        # Classificar tipo de exame US
        df_us['TIPO_EXAME_US'] = df_us.apply(self.classificar_tipo_us, axis=1)

        return df_us

    def apply_dark_theme(self):
        """Aplica tema dark"""
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #0d1117;
                color: #c9d1d9;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
            }
            QPushButton {
                background-color: #21262d;
                border: 1px solid #30363d;
                border-radius: 6px;
                padding: 8px 16px;
                color: #c9d1d9;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #30363d;
                border-color: #58a6ff;
            }
            QPushButton#primaryButton {
                background-color: #238636;
                border-color: #238636;
            }
            QPushButton#primaryButton:hover {
                background-color: #2ea043;
            }
            QPushButton#importButton {
                background-color: #1f6feb;
                border-color: #1f6feb;
            }
            QPushButton#importButton:hover {
                background-color: #388bfd;
            }
            QGroupBox {
                background-color: #161b22;
                border: 1px solid #30363d;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 15px;
                font-weight: bold;
                color: #58a6ff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QTableWidget {
                background-color: #0d1117;
                alternate-background-color: #161b22;
                border: 1px solid #30363d;
                border-radius: 6px;
                gridline-color: #30363d;
            }
            QTableWidget::item {
                padding: 5px;
                color: #c9d1d9;
            }
            QTableWidget::item:selected {
                background-color: #58a6ff;
                color: #ffffff;
            }
            QHeaderView::section {
                background-color: #161b22;
                color: #58a6ff;
                padding: 8px;
                border: none;
                border-right: 1px solid #30363d;
                border-bottom: 1px solid #30363d;
                font-weight: bold;
            }
            QScrollArea {
                border: none;
                background-color: #0d1117;
            }
            QTabWidget::pane {
                border: 1px solid #30363d;
                background-color: #0d1117;
                border-radius: 6px;
            }
            QTabBar::tab {
                background-color: #21262d;
                color: #c9d1d9;
                padding: 10px 20px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-weight: 500;
                font-size: 14px;
            }
            QTabBar::tab:selected {
                background-color: #6f42c1;
                color: white;
            }
            QTabBar::tab:hover {
                background-color: #30363d;
            }
            QSpinBox, QDoubleSpinBox {
                background-color: #21262d;
                border: 1px solid #30363d;
                border-radius: 6px;
                padding: 6px;
                color: #c9d1d9;
            }
            QSpinBox:hover, QDoubleSpinBox:hover {
                border-color: #58a6ff;
            }
            QLabel {
                color: #c9d1d9;
            }
            QComboBox {
                background-color: #21262d;
                border: 1px solid #30363d;
                border-radius: 6px;
                padding: 6px;
                color: #c9d1d9;
            }
            QComboBox:hover {
                border-color: #58a6ff;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox QAbstractItemView {
                background-color: #21262d;
                border: 1px solid #30363d;
                selection-background-color: #58a6ff;
                color: #c9d1d9;
            }
        """)

    def setup_ui(self):
        """Configura interface"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Header
        header = QLabel("🔬 Análise Estratégica de Ultrassonografia")
        header.setStyleSheet("""
            font-size: 24px;
            font-weight: bold;
            color: #6f42c1;
            padding: 10px;
        """)
        main_layout.addWidget(header)

        # Informações do período e dados
        info_group = QGroupBox("📋 Informações dos Dados")
        info_layout = QGridLayout()

        # Período
        periodo_label = QLabel("Período:")
        periodo_value = QLabel(f"{self.start_date.strftime('%d/%m/%Y')} até {self.end_date.strftime('%d/%m/%Y')}")
        periodo_value.setStyleSheet("color: #58a6ff; font-weight: bold;")

        # Total de exames US
        total_label = QLabel("Total de Exames US:")
        total_count = len(self.df_us) if self.df_us is not None else 0
        total_value = QLabel(f"{total_count:,} exames")
        total_value.setStyleSheet("color: #6f42c1; font-weight: bold; font-size: 16px;")

        # Configuração de capacidade
        cap_label = QLabel("Capacidade Diária (exames/dia):")
        self.capacidade_spin = QSpinBox()
        self.capacidade_spin.setRange(1, 500)
        self.capacidade_spin.setValue(80)
        self.capacidade_spin.valueChanged.connect(self.on_capacidade_changed)

        info_layout.addWidget(periodo_label, 0, 0)
        info_layout.addWidget(periodo_value, 0, 1)
        info_layout.addWidget(total_label, 0, 2)
        info_layout.addWidget(total_value, 0, 3)
        info_layout.addWidget(cap_label, 0, 4)
        info_layout.addWidget(self.capacidade_spin, 0, 5)

        info_group.setLayout(info_layout)
        main_layout.addWidget(info_group)

        # Tabs para análises
        self.tabs = QTabWidget()

        # Tab 1: Volumetria Total
        self.tab_volumetria = QWidget()
        self.tab_volumetria_layout = QVBoxLayout(self.tab_volumetria)
        scroll_vol = QScrollArea()
        scroll_vol.setWidgetResizable(True)
        scroll_vol.setFrameShape(QFrame.NoFrame)
        self.volumetria_content = QWidget()
        self.volumetria_content_layout = QVBoxLayout(self.volumetria_content)
        scroll_vol.setWidget(self.volumetria_content)
        self.tab_volumetria_layout.addWidget(scroll_vol)

        # Tab 2: Produtividade por Porta de Entrada
        self.tab_produtividade = QWidget()
        self.tab_produtividade_layout = QVBoxLayout(self.tab_produtividade)
        scroll_prod = QScrollArea()
        scroll_prod.setWidgetResizable(True)
        scroll_prod.setFrameShape(QFrame.NoFrame)
        self.produtividade_content = QWidget()
        self.produtividade_content_layout = QVBoxLayout(self.produtividade_content)
        scroll_prod.setWidget(self.produtividade_content)
        self.tab_produtividade_layout.addWidget(scroll_prod)

        # Tab 3: Distribuição por Tipo de Exame
        self.tab_distribuicao = QWidget()
        self.tab_distribuicao_layout = QVBoxLayout(self.tab_distribuicao)
        scroll_dist = QScrollArea()
        scroll_dist.setWidgetResizable(True)
        scroll_dist.setFrameShape(QFrame.NoFrame)
        self.distribuicao_content = QWidget()
        self.distribuicao_content_layout = QVBoxLayout(self.distribuicao_content)
        scroll_dist.setWidget(self.distribuicao_content)
        self.tab_distribuicao_layout.addWidget(scroll_dist)

        # Tab 4: Demanda vs Capacidade
        self.tab_demanda = QWidget()
        self.tab_demanda_layout = QVBoxLayout(self.tab_demanda)
        scroll_dem = QScrollArea()
        scroll_dem.setWidgetResizable(True)
        scroll_dem.setFrameShape(QFrame.NoFrame)
        self.demanda_content = QWidget()
        self.demanda_content_layout = QVBoxLayout(self.demanda_content)
        scroll_dem.setWidget(self.demanda_content)
        self.tab_demanda_layout.addWidget(scroll_dem)

        self.tabs.addTab(self.tab_volumetria, "📊 Volumetria Total")
        self.tabs.addTab(self.tab_produtividade, "🚪 Produtividade por Porta")
        self.tabs.addTab(self.tab_distribuicao, "🔍 Tipo de Exame")
        self.tabs.addTab(self.tab_demanda, "⚖️ Demanda vs Capacidade")

        main_layout.addWidget(self.tabs, stretch=1)

        # Botões de ação
        button_layout = QHBoxLayout()

        self.export_btn = QPushButton("💾 Exportar Relatório")
        self.export_btn.setObjectName("primaryButton")
        self.export_btn.clicked.connect(self.export_report)
        self.export_btn.setEnabled(self.df_us is not None and len(self.df_us) > 0)

        self.close_btn = QPushButton("✖ Fechar")
        self.close_btn.clicked.connect(self.close)

        button_layout.addWidget(self.export_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.close_btn)

        main_layout.addLayout(button_layout)

    def classificar_tipo_us(self, row):
        """Classifica o tipo de exame de ultrassonografia"""
        # Tenta usar coluna DESCRICAO_PROCEDIMENTO (prioridade) ou variações
        proc = ''
        for col in ['DESCRICAO_PROCEDIMENTO', 'DESCRICAO PROCEDIMENTO', 'DS_PROCEDIMENTO',
                    'PROCEDIMENTO', 'DESCRICAO', 'EXAME', 'NM_PROCEDIMENTO']:
            if col in row.index and pd.notna(row[col]):
                proc = str(row[col]).upper()
                break

        # Classificação por tipo
        if any(x in proc for x in ['OBSTETR', 'GESTAC', 'FETAL', 'MORFOLOG', 'GRAVID']):
            return 'Obstétrico'
        elif any(x in proc for x in ['ABDOM', 'FIGADO', 'HEPAT', 'PANCREA', 'BACO', 'VESICULA', 'RENAL', 'RIM', 'RINS']):
            return 'Abdominal'
        elif any(x in proc for x in ['VASCUL', 'DOPPLER', 'CAROTID', 'VENOSO', 'ARTERIAL', 'VARIZES']):
            return 'Vascular/Doppler'
        elif any(x in proc for x in ['TIREOIDE', 'TIREOID', 'CERVICAL', 'PESCOCO']):
            return 'Tireoide/Cervical'
        elif any(x in proc for x in ['MAMA', 'MAMARIA', 'AXILA']):
            return 'Mama'
        elif any(x in proc for x in ['PROST', 'TRANSRET', 'ESCROTO', 'TESTIC', 'PELVIC', 'UTERO', 'OVARIO', 'TRANSVAG']):
            return 'Pélvico/Urológico'
        elif any(x in proc for x in ['MUSCUL', 'TENDAO', 'ARTICUL', 'OMBRO', 'JOELHO', 'TORNOZELO', 'PUNHO']):
            return 'Musculoesquelético'
        elif any(x in proc for x in ['PARTES MOLES', 'SUBCUTAN', 'HERNIA']):
            return 'Partes Moles'
        else:
            return 'Outros'

    def on_capacidade_changed(self, value):
        """Callback quando capacidade é alterada"""
        self.capacidade_diaria = value
        if self.df_us is not None:
            self.update_demanda_analysis()

    def generate_all_analyses(self):
        """Gera todas as análises"""
        if self.df_us is None or len(self.df_us) == 0:
            return

        self.generate_volumetria_analysis()
        self.generate_produtividade_analysis()
        self.generate_distribuicao_analysis()
        self.generate_demanda_analysis()

    def clear_layout(self, layout):
        """Limpa layout"""
        while layout.count():
            item = layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def generate_volumetria_analysis(self):
        """Gera análise de volumetria total"""
        self.clear_layout(self.volumetria_content_layout)

        # Identificar coluna de data
        date_col = None
        for col in ['DATA_HORA_PRESCRICAO', 'DATA_EXAME', 'DATA_PRESCRICAO', 'DATA']:
            if col in self.df_us.columns:
                date_col = col
                break

        if date_col is None:
            self.volumetria_content_layout.addWidget(
                QLabel("⚠️ Coluna de data não encontrada no arquivo")
            )
            return

        df = self.df_us.copy()
        df['DATA'] = pd.to_datetime(df[date_col]).dt.date
        df['MES_ANO'] = pd.to_datetime(df[date_col]).dt.to_period('M')

        # KPI Cards
        kpi_layout = QHBoxLayout()

        total_exames = len(df)
        media_diaria = df.groupby('DATA').size().mean()
        media_mensal = df.groupby('MES_ANO').size().mean()
        dias_analisados = df['DATA'].nunique()

        card1 = StatCard("Total de Exames US", f"{total_exames:,}",
                        f"{dias_analisados} dias analisados", "#6f42c1")
        card2 = StatCard("Média Diária", f"{media_diaria:.1f}",
                        "exames/dia", "#2196F3")
        card3 = StatCard("Média Mensal", f"{media_mensal:.0f}",
                        "exames/mês", "#00BCD4")

        kpi_layout.addWidget(card1)
        kpi_layout.addWidget(card2)
        kpi_layout.addWidget(card3)

        kpi_widget = QWidget()
        kpi_widget.setLayout(kpi_layout)
        self.volumetria_content_layout.addWidget(kpi_widget)

        # Gráfico de volume mensal
        vol_mensal = df.groupby('MES_ANO').size()

        chart_mensal = MatplotlibChart()
        ax = chart_mensal.figure.add_subplot(111)

        x_labels = [str(p) for p in vol_mensal.index]
        bars = ax.bar(range(len(vol_mensal)), vol_mensal.values,
                     color='#6f42c1', edgecolor='white', linewidth=1.5)

        ax.set_xticks(range(len(vol_mensal)))
        ax.set_xticklabels(x_labels, rotation=45, ha='right', color='#c9d1d9')
        ax.set_ylabel('Quantidade de Exames', color='#c9d1d9', fontweight='bold')
        ax.set_title('Volume Mensal de Exames de Ultrassonografia',
                    color='#6f42c1', fontweight='bold', pad=20)
        ax.tick_params(colors='#c9d1d9')
        ax.set_facecolor('#0d1117')
        ax.grid(True, alpha=0.2, color='#30363d')

        for spine in ax.spines.values():
            spine.set_color('#30363d')

        # Valores nas barras
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{int(height):,}', ha='center', va='bottom',
                   color='#c9d1d9', fontweight='bold', fontsize=9)

        chart_mensal.figure.tight_layout(pad=2.0)
        chart_mensal.canvas.draw()

        container_mensal = QGroupBox("📈 Volume Mensal")
        layout_mensal = QVBoxLayout()
        layout_mensal.addWidget(chart_mensal)
        container_mensal.setLayout(layout_mensal)
        self.volumetria_content_layout.addWidget(container_mensal)

        # Gráfico de volume diário (últimos 30 dias ou todos se menos)
        vol_diario = df.groupby('DATA').size().tail(60)

        chart_diario = MatplotlibChart()
        ax2 = chart_diario.figure.add_subplot(111)

        x_labels_diario = [str(d) for d in vol_diario.index]
        ax2.plot(range(len(vol_diario)), vol_diario.values,
                color='#58a6ff', linewidth=2, marker='o', markersize=4)
        ax2.fill_between(range(len(vol_diario)), vol_diario.values,
                        alpha=0.3, color='#58a6ff')

        # Linha de média
        ax2.axhline(y=vol_diario.mean(), color='#FFC107', linestyle='--',
                   linewidth=2, label=f'Média: {vol_diario.mean():.1f}')

        step = max(1, len(vol_diario) // 10)
        ax2.set_xticks(range(0, len(vol_diario), step))
        ax2.set_xticklabels([x_labels_diario[i] for i in range(0, len(vol_diario), step)],
                           rotation=45, ha='right', color='#c9d1d9', fontsize=8)
        ax2.set_ylabel('Quantidade de Exames', color='#c9d1d9', fontweight='bold')
        ax2.set_title('Volume Diário de Exames (últimos 60 dias)',
                     color='#58a6ff', fontweight='bold', pad=20)
        ax2.tick_params(colors='#c9d1d9')
        ax2.set_facecolor('#0d1117')
        ax2.grid(True, alpha=0.2, color='#30363d')
        ax2.legend(facecolor='#161b22', edgecolor='#30363d', labelcolor='#c9d1d9')

        for spine in ax2.spines.values():
            spine.set_color('#30363d')

        chart_diario.figure.tight_layout(pad=2.0)
        chart_diario.canvas.draw()

        container_diario = QGroupBox("📅 Volume Diário")
        layout_diario = QVBoxLayout()
        layout_diario.addWidget(chart_diario)
        container_diario.setLayout(layout_diario)
        self.volumetria_content_layout.addWidget(container_diario)

        # Tabela resumo mensal
        tabela_group = QGroupBox("📋 Resumo Mensal Detalhado")
        tabela_layout = QVBoxLayout()

        resumo_mensal = df.groupby('MES_ANO').agg({
            'DATA': ['count', 'nunique']
        }).reset_index()
        resumo_mensal.columns = ['Mês', 'Total Exames', 'Dias com Exames']
        resumo_mensal['Média Diária'] = resumo_mensal['Total Exames'] / resumo_mensal['Dias com Exames']

        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(['Mês', 'Total Exames', 'Dias com Exames', 'Média Diária'])
        table.setRowCount(len(resumo_mensal))
        table.setAlternatingRowColors(True)

        for i, row in resumo_mensal.iterrows():
            table.setItem(i, 0, QTableWidgetItem(str(row['Mês'])))
            table.setItem(i, 1, QTableWidgetItem(f"{int(row['Total Exames']):,}"))
            table.setItem(i, 2, QTableWidgetItem(f"{int(row['Dias com Exames'])}"))
            table.setItem(i, 3, QTableWidgetItem(f"{row['Média Diária']:.1f}"))

        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setMinimumHeight(200)
        tabela_layout.addWidget(table)
        tabela_group.setLayout(tabela_layout)
        self.volumetria_content_layout.addWidget(tabela_group)

    def generate_produtividade_analysis(self):
        """Gera análise de produtividade por porta de entrada"""
        self.clear_layout(self.produtividade_content_layout)

        if 'TIPO_ATENDIMENTO' not in self.df_us.columns:
            self.produtividade_content_layout.addWidget(
                QLabel("⚠️ Coluna TIPO_ATENDIMENTO não encontrada")
            )
            return

        df = self.df_us.copy()

        # Análise por tipo de atendimento
        analise_tipo = df.groupby('TIPO_ATENDIMENTO').size().reset_index(name='Total')
        analise_tipo = analise_tipo.sort_values('Total', ascending=False)
        analise_tipo['Percentual'] = (analise_tipo['Total'] / analise_tipo['Total'].sum() * 100)

        # KPI Cards por porta
        kpi_layout = QHBoxLayout()
        cores = ['#FF6B6B', '#51CF66', '#4E8FDF', '#FFC107', '#9C27B0']

        for i, (_, row) in enumerate(analise_tipo.head(4).iterrows()):
            cor = cores[i % len(cores)]
            card = StatCard(
                row['TIPO_ATENDIMENTO'][:20],
                f"{int(row['Total']):,}",
                f"{row['Percentual']:.1f}% do total",
                cor
            )
            kpi_layout.addWidget(card)

        kpi_widget = QWidget()
        kpi_widget.setLayout(kpi_layout)
        self.produtividade_content_layout.addWidget(kpi_widget)

        # Gráfico de pizza
        chart_pizza = MatplotlibChart()
        ax = chart_pizza.figure.add_subplot(111)

        colors = ['#FF6B6B', '#51CF66', '#4E8FDF', '#FFC107', '#9C27B0', '#00BCD4', '#FF9800']
        explode = [0.02] * len(analise_tipo)

        wedges, texts, autotexts = ax.pie(
            analise_tipo['Total'].values,
            labels=analise_tipo['TIPO_ATENDIMENTO'].values,
            autopct='%1.1f%%',
            colors=colors[:len(analise_tipo)],
            explode=explode,
            textprops={'color': '#c9d1d9', 'fontsize': 10}
        )

        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')

        ax.set_title('Distribuição por Porta de Entrada',
                    color='#6f42c1', fontweight='bold', pad=20)

        chart_pizza.figure.tight_layout(pad=2.0)
        chart_pizza.canvas.draw()

        container_pizza = QGroupBox("🥧 Distribuição Percentual")
        layout_pizza = QVBoxLayout()
        layout_pizza.addWidget(chart_pizza)
        container_pizza.setLayout(layout_pizza)
        self.produtividade_content_layout.addWidget(container_pizza)

        # Gráfico de barras
        chart_barras = MatplotlibChart()
        ax2 = chart_barras.figure.add_subplot(111)

        bars = ax2.barh(range(len(analise_tipo)), analise_tipo['Total'].values,
                       color=colors[:len(analise_tipo)], edgecolor='white', linewidth=1.5)

        ax2.set_yticks(range(len(analise_tipo)))
        ax2.set_yticklabels(analise_tipo['TIPO_ATENDIMENTO'].values, color='#c9d1d9')
        ax2.set_xlabel('Quantidade de Exames', color='#c9d1d9', fontweight='bold')
        ax2.set_title('Volume por Porta de Entrada', color='#58a6ff', fontweight='bold', pad=20)
        ax2.tick_params(colors='#c9d1d9')
        ax2.set_facecolor('#0d1117')
        ax2.grid(True, alpha=0.2, axis='x', color='#30363d')

        for spine in ax2.spines.values():
            spine.set_color('#30363d')

        for bar in bars:
            width = bar.get_width()
            ax2.text(width + 50, bar.get_y() + bar.get_height()/2.,
                    f'{int(width):,}', ha='left', va='center',
                    color='#c9d1d9', fontweight='bold')

        chart_barras.figure.tight_layout(pad=2.0)
        chart_barras.canvas.draw()

        container_barras = QGroupBox("📊 Volume Absoluto")
        layout_barras = QVBoxLayout()
        layout_barras.addWidget(chart_barras)
        container_barras.setLayout(layout_barras)
        self.produtividade_content_layout.addWidget(container_barras)

        # Tabela detalhada
        tabela_group = QGroupBox("📋 Tabela Detalhada por Porta de Entrada")
        tabela_layout = QVBoxLayout()

        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(['Porta de Entrada', 'Total Exames', 'Percentual'])
        table.setRowCount(len(analise_tipo))
        table.setAlternatingRowColors(True)

        for i, (_, row) in enumerate(analise_tipo.iterrows()):
            table.setItem(i, 0, QTableWidgetItem(str(row['TIPO_ATENDIMENTO'])))
            table.setItem(i, 1, QTableWidgetItem(f"{int(row['Total']):,}"))
            table.setItem(i, 2, QTableWidgetItem(f"{row['Percentual']:.1f}%"))

        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setMinimumHeight(200)
        tabela_layout.addWidget(table)
        tabela_group.setLayout(tabela_layout)
        self.produtividade_content_layout.addWidget(tabela_group)

    def generate_distribuicao_analysis(self):
        """Gera análise de distribuição por tipo de exame US"""
        self.clear_layout(self.distribuicao_content_layout)

        df = self.df_us.copy()

        # Análise por tipo de exame US
        analise_tipo_us = df.groupby('TIPO_EXAME_US').size().reset_index(name='Total')
        analise_tipo_us = analise_tipo_us.sort_values('Total', ascending=False)
        analise_tipo_us['Percentual'] = (analise_tipo_us['Total'] / analise_tipo_us['Total'].sum() * 100)

        # KPI Cards
        kpi_layout = QHBoxLayout()
        cores = {
            'Abdominal': '#4CAF50',
            'Obstétrico': '#E91E63',
            'Vascular/Doppler': '#2196F3',
            'Tireoide/Cervical': '#9C27B0',
            'Mama': '#FF9800',
            'Pélvico/Urológico': '#00BCD4',
            'Musculoesquelético': '#795548',
            'Partes Moles': '#607D8B',
            'Outros': '#9E9E9E'
        }

        for _, row in analise_tipo_us.head(4).iterrows():
            cor = cores.get(row['TIPO_EXAME_US'], '#6f42c1')
            card = StatCard(
                row['TIPO_EXAME_US'],
                f"{int(row['Total']):,}",
                f"{row['Percentual']:.1f}%",
                cor
            )
            kpi_layout.addWidget(card)

        kpi_widget = QWidget()
        kpi_widget.setLayout(kpi_layout)
        self.distribuicao_content_layout.addWidget(kpi_widget)

        # Gráfico de pizza
        chart_pizza = MatplotlibChart()
        ax = chart_pizza.figure.add_subplot(111)

        colors_list = [cores.get(t, '#6f42c1') for t in analise_tipo_us['TIPO_EXAME_US']]

        wedges, texts, autotexts = ax.pie(
            analise_tipo_us['Total'].values,
            labels=analise_tipo_us['TIPO_EXAME_US'].values,
            autopct='%1.1f%%',
            colors=colors_list,
            textprops={'color': '#c9d1d9', 'fontsize': 9}
        )

        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(8)

        ax.set_title('Distribuição por Tipo de Exame de Ultrassonografia',
                    color='#6f42c1', fontweight='bold', pad=20)

        chart_pizza.figure.tight_layout(pad=2.0)
        chart_pizza.canvas.draw()

        container_pizza = QGroupBox("🥧 Distribuição por Tipo de Exame")
        layout_pizza = QVBoxLayout()
        layout_pizza.addWidget(chart_pizza)
        container_pizza.setLayout(layout_pizza)
        self.distribuicao_content_layout.addWidget(container_pizza)

        # Gráfico de barras horizontal
        chart_barras = MatplotlibChart()
        ax2 = chart_barras.figure.add_subplot(111)

        y_pos = range(len(analise_tipo_us))
        bars = ax2.barh(y_pos, analise_tipo_us['Total'].values,
                       color=colors_list, edgecolor='white', linewidth=1.5)

        ax2.set_yticks(y_pos)
        ax2.set_yticklabels(analise_tipo_us['TIPO_EXAME_US'].values, color='#c9d1d9')
        ax2.set_xlabel('Quantidade de Exames', color='#c9d1d9', fontweight='bold')
        ax2.set_title('Volume por Tipo de Exame US', color='#58a6ff', fontweight='bold', pad=20)
        ax2.tick_params(colors='#c9d1d9')
        ax2.set_facecolor('#0d1117')
        ax2.grid(True, alpha=0.2, axis='x', color='#30363d')

        for spine in ax2.spines.values():
            spine.set_color('#30363d')

        for bar in bars:
            width = bar.get_width()
            ax2.text(width + 20, bar.get_y() + bar.get_height()/2.,
                    f'{int(width):,}', ha='left', va='center',
                    color='#c9d1d9', fontweight='bold', fontsize=9)

        chart_barras.figure.tight_layout(pad=2.0)
        chart_barras.canvas.draw()

        container_barras = QGroupBox("📊 Volume por Tipo")
        layout_barras = QVBoxLayout()
        layout_barras.addWidget(chart_barras)
        container_barras.setLayout(layout_barras)
        self.distribuicao_content_layout.addWidget(container_barras)

        # Tabela detalhada
        tabela_group = QGroupBox("📋 Tabela Detalhada por Tipo de Exame")
        tabela_layout = QVBoxLayout()

        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(['Tipo de Exame', 'Total', 'Percentual'])
        table.setRowCount(len(analise_tipo_us))
        table.setAlternatingRowColors(True)

        for i, (_, row) in enumerate(analise_tipo_us.iterrows()):
            table.setItem(i, 0, QTableWidgetItem(str(row['TIPO_EXAME_US'])))
            table.setItem(i, 1, QTableWidgetItem(f"{int(row['Total']):,}"))
            table.setItem(i, 2, QTableWidgetItem(f"{row['Percentual']:.1f}%"))

        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setMinimumHeight(250)
        tabela_layout.addWidget(table)
        tabela_group.setLayout(tabela_layout)
        self.distribuicao_content_layout.addWidget(tabela_group)

    def generate_demanda_analysis(self):
        """Gera análise de demanda vs capacidade"""
        self.update_demanda_analysis()

    def update_demanda_analysis(self):
        """Atualiza análise de demanda vs capacidade (apenas porta EXTERNA)"""
        self.clear_layout(self.demanda_content_layout)

        if self.df_us is None:
            return

        # Identificar coluna de data
        date_col = None
        for col in ['DATA_HORA_PRESCRICAO', 'DATA_EXAME', 'DATA_PRESCRICAO', 'DATA']:
            if col in self.df_us.columns:
                date_col = col
                break

        if date_col is None:
            self.demanda_content_layout.addWidget(
                QLabel("⚠️ Coluna de data não encontrada")
            )
            return

        # Filtrar apenas porta EXTERNA
        df = self.df_us.copy()
        if 'TIPO_ATENDIMENTO' in df.columns:
            df = df[df['TIPO_ATENDIMENTO'].str.upper().str.contains('EXTERNO', na=False)]

        if len(df) == 0:
            self.demanda_content_layout.addWidget(
                QLabel("⚠️ Nenhum exame de porta EXTERNA encontrado")
            )
            return

        df['DATA'] = pd.to_datetime(df[date_col]).dt.date

        # Volume diário
        vol_diario = df.groupby('DATA').size().reset_index(name='Demanda')
        vol_diario['Capacidade'] = self.capacidade_diaria
        vol_diario['Gap'] = vol_diario['Demanda'] - vol_diario['Capacidade']
        vol_diario['Status'] = vol_diario['Gap'].apply(
            lambda x: 'Acima' if x > 0 else ('Abaixo' if x < 0 else 'Igual')
        )

        # Estatísticas
        media_demanda = vol_diario['Demanda'].mean()
        max_demanda = vol_diario['Demanda'].max()
        min_demanda = vol_diario['Demanda'].min()
        dias_acima = (vol_diario['Gap'] > 0).sum()
        dias_abaixo = (vol_diario['Gap'] <= 0).sum()
        utilizacao = (media_demanda / self.capacidade_diaria * 100)

        # KPI Cards
        kpi_layout = QHBoxLayout()

        card1 = StatCard("Capacidade Diária", f"{self.capacidade_diaria}",
                        "exames/dia configurados", "#6f42c1")
        card2 = StatCard("Demanda Média", f"{media_demanda:.1f}",
                        f"Min: {min_demanda} / Max: {max_demanda}", "#2196F3")
        card3 = StatCard("Taxa Utilização", f"{utilizacao:.1f}%",
                        "demanda/capacidade",
                        "#4CAF50" if utilizacao <= 100 else "#F44336")
        card4 = StatCard("Dias Acima Capacidade", f"{dias_acima}",
                        f"de {len(vol_diario)} dias",
                        "#F44336" if dias_acima > dias_abaixo else "#4CAF50")

        kpi_layout.addWidget(card1)
        kpi_layout.addWidget(card2)
        kpi_layout.addWidget(card3)
        kpi_layout.addWidget(card4)

        kpi_widget = QWidget()
        kpi_widget.setLayout(kpi_layout)
        self.demanda_content_layout.addWidget(kpi_widget)

        # Gráfico de linha: Demanda vs Capacidade
        chart_linha = MatplotlibChart()
        ax = chart_linha.figure.add_subplot(111)

        x = range(len(vol_diario))
        ax.plot(x, vol_diario['Demanda'].values, color='#58a6ff',
               linewidth=2, label='Demanda Real', marker='o', markersize=3)
        ax.axhline(y=self.capacidade_diaria, color='#F44336',
                  linestyle='--', linewidth=2, label=f'Capacidade ({self.capacidade_diaria})')
        ax.fill_between(x, vol_diario['Demanda'].values, self.capacidade_diaria,
                       where=(vol_diario['Demanda'].values > self.capacidade_diaria),
                       alpha=0.3, color='#F44336', label='Acima da capacidade')
        ax.fill_between(x, vol_diario['Demanda'].values, self.capacidade_diaria,
                       where=(vol_diario['Demanda'].values <= self.capacidade_diaria),
                       alpha=0.3, color='#4CAF50', label='Dentro da capacidade')

        step = max(1, len(vol_diario) // 10)
        ax.set_xticks(range(0, len(vol_diario), step))
        x_labels = [str(vol_diario.iloc[i]['DATA']) for i in range(0, len(vol_diario), step)]
        ax.set_xticklabels(x_labels, rotation=45, ha='right', color='#c9d1d9', fontsize=8)
        ax.set_ylabel('Quantidade de Exames', color='#c9d1d9', fontweight='bold')
        ax.set_title('Demanda Diária vs Capacidade - Porta EXTERNA',
                    color='#6f42c1', fontweight='bold', pad=20)
        ax.tick_params(colors='#c9d1d9')
        ax.set_facecolor('#0d1117')
        ax.grid(True, alpha=0.2, color='#30363d')
        ax.legend(facecolor='#161b22', edgecolor='#30363d', labelcolor='#c9d1d9',
                 loc='upper right', fontsize=9)

        for spine in ax.spines.values():
            spine.set_color('#30363d')

        chart_linha.figure.tight_layout(pad=2.0)
        chart_linha.canvas.draw()

        container_linha = QGroupBox("📈 Demanda vs Capacidade - Porta EXTERNA (Série Temporal)")
        layout_linha = QVBoxLayout()
        layout_linha.addWidget(chart_linha)
        container_linha.setLayout(layout_linha)
        self.demanda_content_layout.addWidget(container_linha)

        # Histograma de demanda
        chart_hist = MatplotlibChart()
        ax2 = chart_hist.figure.add_subplot(111)

        n, bins, patches = ax2.hist(vol_diario['Demanda'].values, bins=20,
                                   color='#6f42c1', edgecolor='white', alpha=0.7)

        # Colorir barras baseado na capacidade
        for patch, left_edge in zip(patches, bins[:-1]):
            if left_edge > self.capacidade_diaria:
                patch.set_facecolor('#F44336')
            else:
                patch.set_facecolor('#4CAF50')

        ax2.axvline(x=self.capacidade_diaria, color='#FFC107', linestyle='--',
                   linewidth=2, label=f'Capacidade ({self.capacidade_diaria})')
        ax2.axvline(x=media_demanda, color='#58a6ff', linestyle='-',
                   linewidth=2, label=f'Média ({media_demanda:.1f})')

        ax2.set_xlabel('Demanda Diária (exames)', color='#c9d1d9', fontweight='bold')
        ax2.set_ylabel('Frequência (dias)', color='#c9d1d9', fontweight='bold')
        ax2.set_title('Distribuição da Demanda Diária - Porta EXTERNA',
                     color='#58a6ff', fontweight='bold', pad=20)
        ax2.tick_params(colors='#c9d1d9')
        ax2.set_facecolor('#0d1117')
        ax2.grid(True, alpha=0.2, color='#30363d')
        ax2.legend(facecolor='#161b22', edgecolor='#30363d', labelcolor='#c9d1d9')

        for spine in ax2.spines.values():
            spine.set_color('#30363d')

        chart_hist.figure.tight_layout(pad=2.0)
        chart_hist.canvas.draw()

        container_hist = QGroupBox("📊 Histograma de Demanda")
        layout_hist = QVBoxLayout()
        layout_hist.addWidget(chart_hist)
        container_hist.setLayout(layout_hist)
        self.demanda_content_layout.addWidget(container_hist)

        # Tabela de dias críticos (acima da capacidade)
        dias_criticos = vol_diario[vol_diario['Gap'] > 0].sort_values('Gap', ascending=False)

        if len(dias_criticos) > 0:
            tabela_group = QGroupBox(f"⚠️ Dias Críticos (Demanda > Capacidade) - {len(dias_criticos)} dias")
            tabela_layout = QVBoxLayout()

            table = QTableWidget()
            table.setColumnCount(4)
            table.setHorizontalHeaderLabels(['Data', 'Demanda', 'Capacidade', 'Excesso'])
            table.setRowCount(min(20, len(dias_criticos)))
            table.setAlternatingRowColors(True)

            for i, (_, row) in enumerate(dias_criticos.head(20).iterrows()):
                table.setItem(i, 0, QTableWidgetItem(str(row['DATA'])))
                table.setItem(i, 1, QTableWidgetItem(f"{int(row['Demanda'])}"))
                table.setItem(i, 2, QTableWidgetItem(f"{int(row['Capacidade'])}"))

                gap_item = QTableWidgetItem(f"+{int(row['Gap'])}")
                gap_item.setForeground(QColor('#F44336'))
                table.setItem(i, 3, gap_item)

            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.setMinimumHeight(200)
            tabela_layout.addWidget(table)
            tabela_group.setLayout(tabela_layout)
            self.demanda_content_layout.addWidget(tabela_group)

    def export_report(self):
        """Exporta relatório para Excel"""
        if self.df_us is None:
            QMessageBox.warning(self, "Aviso", "Nenhum dado carregado para exportar.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar Relatório",
            build_runtime_file_path(
                f"relatorio_us_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            ),
            "Excel Files (*.xlsx)"
        )

        if not file_path:
            return

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Dados brutos
                self.df_us.to_excel(writer, sheet_name='Dados_Brutos', index=False)

                # Resumo por tipo de exame
                resumo_tipo = self.df_us.groupby('TIPO_EXAME_US').size().reset_index(name='Total')
                resumo_tipo.to_excel(writer, sheet_name='Por_Tipo_Exame', index=False)

                # Resumo por porta de entrada
                if 'TIPO_ATENDIMENTO' in self.df_us.columns:
                    resumo_porta = self.df_us.groupby('TIPO_ATENDIMENTO').size().reset_index(name='Total')
                    resumo_porta.to_excel(writer, sheet_name='Por_Porta_Entrada', index=False)

                # Resumo por convênio (coluna CONVENIO)
                convenio_col = None
                for col in ['CONVENIO', 'CONVÊNIO', 'convenio', 'Convênio']:
                    if col in self.df_us.columns:
                        convenio_col = col
                        break

                if convenio_col is not None:
                    convenio_series = self.df_us[convenio_col].astype(str).str.strip()
                    convenio_series = convenio_series.replace({'nan': '', 'NaT': '', 'None': '', 'NONE': ''})
                    df_convenio = self.df_us[convenio_series != ''].copy()
                    if len(df_convenio) > 0:
                        df_convenio['CONVENIO_LIMPO'] = convenio_series[convenio_series != '']
                        resumo_convenio = (
                            df_convenio.groupby('CONVENIO_LIMPO').size()
                            .reset_index(name='Total')
                            .sort_values('Total', ascending=False)
                        )
                        resumo_convenio['Percentual'] = (
                            resumo_convenio['Total'] / resumo_convenio['Total'].sum() * 100
                        )
                        resumo_convenio.to_excel(writer, sheet_name='Por_Convenio', index=False)

            QMessageBox.information(self, "Sucesso", f"Relatório exportado com sucesso!\n\n{file_path}")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar relatório:\n{str(e)}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.base_title = "CDI - Análise de Tempo de Atendimento"
        self.setWindowTitle(self.base_title)
        self.setGeometry(100, 100, 1600, 1000)

        # Dados
        self.df = None
        self.df_original = None  # Dados originais sem filtro de hospital
        self.resultado = None
        self.resultado_original = None  # Resultado original sem filtro de hospital
        self.current_file = None
        self.selected_hospital = "Todas as Unidades"  # Hospital selecionado

        # Setup UI
        self.setup_dark_theme()
        self.setup_ui()
        self.setup_menu_bar()

    def setup_dark_theme(self):
        """Configura tema dark moderno"""
        dark_stylesheet = """
            QMainWindow {
                background-color: #0d1117;
            }
            QWidget {
                background-color: #0d1117;
                color: #c9d1d9;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
            }
            QPushButton {
                background-color: #21262d;
                border: 1px solid #30363d;
                border-radius: 6px;
                padding: 8px 16px;
                color: #c9d1d9;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #30363d;
                border-color: #58a6ff;
            }
            QPushButton:pressed {
                background-color: #161b22;
            }
            QPushButton#primaryButton {
                background-color: #238636;
                border-color: #238636;
            }
            QPushButton#primaryButton:hover {
                background-color: #2ea043;
            }
            QLabel {
                color: #c9d1d9;
            }
            QLabel#titleLabel {
                font-size: 24px;
                font-weight: bold;
                color: #58a6ff;
            }
            QGroupBox {
                background-color: #161b22;
                border: 1px solid #30363d;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 15px;
                font-weight: bold;
            }
            QGroupBox::title {
                color: #58a6ff;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QDateEdit {
                background-color: #21262d;
                border: 1px solid #30363d;
                border-radius: 6px;
                padding: 6px;
                color: #c9d1d9;
            }
            QDateEdit:hover {
                border-color: #58a6ff;
            }
            QDateEdit::drop-down {
                border: none;
                width: 20px;
            }
            QDateEdit::down-arrow {
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 6px solid #c9d1d9;
            }
            QCalendarWidget {
                background-color: #161b22;
            }
            QCalendarWidget QToolButton {
                color: #c9d1d9;
                background-color: #21262d;
            }
            QCalendarWidget QMenu {
                background-color: #21262d;
                color: #c9d1d9;
            }
            QCalendarWidget QSpinBox {
                background-color: #21262d;
                color: #c9d1d9;
            }
            QCalendarWidget QTableView {
                background-color: #161b22;
                selection-background-color: #58a6ff;
            }
            QTableWidget {
                background-color: #0d1117;
                alternate-background-color: #161b22;
                border: 1px solid #30363d;
                border-radius: 6px;
                gridline-color: #30363d;
            }
            QTableWidget::item {
                padding: 5px;
                color: #c9d1d9;
            }
            QTableWidget::item:selected {
                background-color: #58a6ff;
                color: #ffffff;
            }
            QHeaderView::section {
                background-color: #161b22;
                color: #58a6ff;
                padding: 8px;
                border: none;
                border-right: 1px solid #30363d;
                border-bottom: 1px solid #30363d;
                font-weight: bold;
            }
            QScrollBar:vertical {
                background: #161b22;
                width: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background: #30363d;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical:hover {
                background: #484f58;
            }
            QScrollBar:horizontal {
                background: #161b22;
                height: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:horizontal {
                background: #30363d;
                border-radius: 6px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #484f58;
            }
            QTabWidget::pane {
                border: 1px solid #30363d;
                background-color: #0d1117;
                border-radius: 6px;
            }
            QTabBar::tab {
                background-color: #21262d;
                color: #c9d1d9;
                padding: 10px 20px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-weight: 500;
                font-size: 14px;
            }
            QTabBar::tab:selected {
                background-color: #238636;
                color: white;
            }
            QTabBar::tab:hover {
                background-color: #30363d;
            }
            QComboBox {
                background-color: #21262d;
                border: 1px solid #30363d;
                border-radius: 6px;
                padding: 6px;
                color: #c9d1d9;
            }
            QComboBox:hover {
                border-color: #58a6ff;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 6px solid #c9d1d9;
            }
            QComboBox QAbstractItemView {
                background-color: #21262d;
                border: 1px solid #30363d;
                selection-background-color: #58a6ff;
                color: #c9d1d9;
            }
        """
        self.setStyleSheet(dark_stylesheet)

    def setup_ui(self):
        """Configura interface"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Header
        header_layout = QHBoxLayout()
        title = QLabel("🏥 Centro de Diagnóstico por Imagem - Análise de Tempo de Atendimento")
        title.setObjectName("titleLabel")
        header_layout.addWidget(title)
        header_layout.addStretch()
        main_layout.addLayout(header_layout)

        # Controles
        controls_group = QGroupBox("Configurações da Análise")
        controls_layout = QGridLayout()

        # Seleção de arquivo
        file_label = QLabel("Arquivo:")
        self.file_path_label = QLabel("Nenhum arquivo selecionado")
        self.file_path_label.setStyleSheet("color: #8b949e; font-style: italic;")
        self.select_file_btn = QPushButton("📁 Selecionar Arquivo Excel")
        self.select_file_btn.setObjectName("primaryButton")
        self.select_file_btn.clicked.connect(self.select_file)

        # Período
        period_label = QLabel("Período:")
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate().addMonths(-1))
        self.start_date.setDisplayFormat("dd/MM/yyyy")

        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())
        self.end_date.setDisplayFormat("dd/MM/yyyy")

        # Botão analisar
        self.analyze_btn = QPushButton("📊 Analisar Dados")
        self.analyze_btn.setObjectName("primaryButton")
        self.analyze_btn.clicked.connect(self.analyze_data)
        self.analyze_btn.setEnabled(False)

        controls_layout.addWidget(file_label, 0, 0)
        controls_layout.addWidget(self.file_path_label, 0, 1, 1, 2)
        controls_layout.addWidget(self.select_file_btn, 0, 3)
        controls_layout.addWidget(period_label, 1, 0)
        controls_layout.addWidget(QLabel("De:"), 1, 1)
        controls_layout.addWidget(self.start_date, 1, 2)
        controls_layout.addWidget(QLabel("Até:"), 1, 3)
        controls_layout.addWidget(self.end_date, 1, 4)
        controls_layout.addWidget(self.analyze_btn, 1, 5)

        controls_group.setLayout(controls_layout)
        main_layout.addWidget(controls_group)

        # Configuração de IA
        ai_group = QGroupBox("🤖 Análise Estratégica com IA")
        ai_layout = QGridLayout()

        # Tipo de API
        ai_label = QLabel("API:")
        self.ai_type_combo = QComboBox()
        self.ai_type_combo.addItems(["OpenAI (ChatGPT)", "LM Studio (Local)"])

        # API Key / URL
        self.api_key_label = QLabel("API Key:")
        self.api_key_input = QLineEdit()
        self.api_key_input.setPlaceholderText("sk-...")
        self.api_key_input.setEchoMode(QLineEdit.Password)

        self.api_url_label = QLabel("URL:")
        self.api_url_input = QLineEdit()
        self.api_url_input.setPlaceholderText("http://localhost:1234/v1")
        self.api_url_input.setText("http://localhost:1234/v1")  # Valor padrão
        self.api_url_input.setVisible(False)

        # Alternar entre API Key e URL
        def on_api_type_changed():
            is_lm_studio = "LM Studio" in self.ai_type_combo.currentText()
            self.api_key_label.setVisible(not is_lm_studio)
            self.api_key_input.setVisible(not is_lm_studio)
            self.api_url_label.setVisible(is_lm_studio)
            self.api_url_input.setVisible(is_lm_studio)

        self.ai_type_combo.currentTextChanged.connect(on_api_type_changed)

        # Botões
        self.search_patient_btn = QPushButton("🔍 Buscar Paciente")
        self.search_patient_btn.setObjectName("primaryButton")
        self.search_patient_btn.clicked.connect(self.open_patient_search)
        self.search_patient_btn.setEnabled(False)

        self.open_dashboard_btn = QPushButton("📊 Abrir Dashboard Completo")
        self.open_dashboard_btn.setObjectName("primaryButton")
        self.open_dashboard_btn.clicked.connect(self.open_dashboard_window)
        self.open_dashboard_btn.setEnabled(False)

        self.generate_ai_btn = QPushButton("🧠 Gerar Análise Estratégica")
        self.generate_ai_btn.setObjectName("primaryButton")
        self.generate_ai_btn.clicked.connect(self.generate_ai_analysis)
        self.generate_ai_btn.setEnabled(False)

        self.us_analysis_btn = QPushButton("🔬 Análise Ultrassonografia")
        self.us_analysis_btn.setObjectName("primaryButton")
        self.us_analysis_btn.setStyleSheet("""
            QPushButton {
                background-color: #6f42c1;
                border-color: #6f42c1;
            }
            QPushButton:hover {
                background-color: #8b5cf6;
            }
        """)
        self.us_analysis_btn.clicked.connect(self.open_us_analysis)

        ai_layout.addWidget(ai_label, 0, 0)
        ai_layout.addWidget(self.ai_type_combo, 0, 1)
        ai_layout.addWidget(self.api_key_label, 0, 2)
        ai_layout.addWidget(self.api_key_input, 0, 3)
        ai_layout.addWidget(self.api_url_label, 0, 2)
        ai_layout.addWidget(self.api_url_input, 0, 3)
        ai_layout.addWidget(self.search_patient_btn, 0, 4)
        ai_layout.addWidget(self.open_dashboard_btn, 0, 5)
        ai_layout.addWidget(self.generate_ai_btn, 0, 6)
        ai_layout.addWidget(self.us_analysis_btn, 0, 7)

        ai_group.setLayout(ai_layout)
        main_layout.addWidget(ai_group)

        # Tabs principais (SLA + Convênios)
        self.main_tabs = QTabWidget()

        # Tab SLA
        self.tab_sla = QWidget()
        self.tab_sla_layout = QVBoxLayout(self.tab_sla)
        scroll_sla = QScrollArea()
        scroll_sla.setWidgetResizable(True)
        scroll_sla.setFrameShape(QFrame.NoFrame)

        self.sla_content = QWidget()
        self.sla_content_layout = QVBoxLayout(self.sla_content)
        self.sla_content_layout.setSpacing(15)
        self.sla_content_layout.setContentsMargins(0, 0, 0, 0)

        # KPI Cards
        self.kpi_layout = QVBoxLayout()
        self.kpi_layout.setSpacing(10)
        self.sla_content_layout.addLayout(self.kpi_layout)

        # Dashboard com gráficos e tabelas
        self.dashboard_widget = QWidget()
        self.dashboard_layout = QVBoxLayout(self.dashboard_widget)
        self.sla_content_layout.addWidget(self.dashboard_widget)

        scroll_sla.setWidget(self.sla_content)
        self.tab_sla_layout.addWidget(scroll_sla)

        # Tab Convênios
        self.tab_convenios_main = QWidget()
        self.tab_convenios_main_layout = QVBoxLayout(self.tab_convenios_main)
        scroll_conv = QScrollArea()
        scroll_conv.setWidgetResizable(True)
        scroll_conv.setFrameShape(QFrame.NoFrame)

        self.convenio_content = QWidget()
        self.convenio_layout = QVBoxLayout(self.convenio_content)
        self.convenio_layout.setSpacing(15)
        self.convenio_layout.setContentsMargins(0, 0, 0, 0)
        scroll_conv.setWidget(self.convenio_content)
        self.tab_convenios_main_layout.addWidget(scroll_conv)

        self.main_tabs.addTab(self.tab_sla, "📊 SLA e Operacional")
        self.main_tabs.addTab(self.tab_convenios_main, "🏥 Convênios")

        main_layout.addWidget(self.main_tabs, stretch=1)

    def setup_menu_bar(self):
        """Configura barra de menu com seletor de hospital"""
        menu_bar = self.menuBar()

        # Menu Unidade
        unidade_menu = menu_bar.addMenu("🏥 Unidade")

        # Criar widget para o combobox
        unidade_widget = QWidget()
        unidade_layout = QHBoxLayout(unidade_widget)
        unidade_layout.setContentsMargins(10, 5, 10, 5)

        unidade_label = QLabel("Selecionar Hospital:")
        self.hospital_combo = QComboBox()
        self.hospital_combo.addItem("Todas as Unidades")
        self.hospital_combo.setMinimumWidth(300)
        self.hospital_combo.currentTextChanged.connect(self.on_hospital_changed)
        self.hospital_combo.setEnabled(False)  # Desabilitado até carregar dados

        unidade_layout.addWidget(unidade_label)
        unidade_layout.addWidget(self.hospital_combo)

        # Criar action e adicionar widget
        widget_action = QWidgetAction(unidade_menu)
        widget_action.setDefaultWidget(unidade_widget)
        unidade_menu.addAction(widget_action)

    def on_hospital_changed(self, hospital_name):
        """Callback quando hospital é alterado"""
        self.selected_hospital = hospital_name

        # Atualizar título da janela
        if hospital_name == "Todas as Unidades":
            self.setWindowTitle(self.base_title)
        else:
            self.setWindowTitle(f"{self.base_title} - {hospital_name}")

        # Se já temos dados carregados, reprocessar com filtro
        if self.resultado_original is not None:
            self.filter_data_by_hospital()

    def filter_data_by_hospital(self):
        """Filtra dados pelo hospital selecionado"""
        if self.resultado_original is None or self.df_original is None:
            return

        # Se "Todas as Unidades", usar dados originais
        if self.selected_hospital == "Todas as Unidades":
            self.resultado = self.resultado_original
            self.df = self.df_original
        else:
            # Filtrar dataframe pelo hospital
            df_filtered = self.df_original[
                self.df_original['UNIDADE'].str.strip() == self.selected_hospital
            ].copy()

            # Recalcular estatísticas com dados filtrados
            self.df = df_filtered
            self.resultado = self.recalculate_stats(df_filtered)

        # Atualizar visualizações
        self.update_kpi_cards(
            self.resultado['stats_gerais'],
            self.resultado['stats_entrega'],
            self.df
        )
        self.create_dashboard(self.resultado)

    def recalculate_stats(self, df):
        """Recalcula estatísticas para o dataframe filtrado"""
        # Filtrar dados de entrega válidos
        df_entrega = df[df['TEMPO_ENTREGA_MIN'].notna()].copy()

        # Estatísticas gerais - ATENDIMENTO
        stats_gerais = {
            'total_exames': len(df),
            'tempo_medio': df['TEMPO_ATENDIMENTO_MIN'].mean() if len(df) > 0 else 0,
            'tempo_mediano': df['TEMPO_ATENDIMENTO_MIN'].median() if len(df) > 0 else 0,
            'tempo_min': df['TEMPO_ATENDIMENTO_MIN'].min() if len(df) > 0 else 0,
            'tempo_max': df['TEMPO_ATENDIMENTO_MIN'].max() if len(df) > 0 else 0,
            'dentro_sla': df['DENTRO_SLA'].sum() if len(df) > 0 else 0,
            'fora_sla': (~df['DENTRO_SLA']).sum() if len(df) > 0 else 0,
            'percentual_sla': (df['DENTRO_SLA'].sum() / len(df) * 100) if len(df) > 0 else 0
        }

        # Estatísticas gerais - ENTREGA
        stats_entrega = {
            'total_exames': len(df_entrega),
            'tempo_medio_entrega': df_entrega['TEMPO_ENTREGA_MIN'].mean() / 60 if len(df_entrega) > 0 else 0,
            'dias_uteis_medio': df_entrega['DIAS_UTEIS_ENTREGA'].mean() if len(df_entrega) > 0 else 0,
            'dentro_sla_entrega': df_entrega['DENTRO_SLA_ENTREGA'].sum() if len(df_entrega) > 0 else 0,
            'fora_sla_entrega': (~df_entrega['DENTRO_SLA_ENTREGA']).sum() if len(df_entrega) > 0 else 0,
            'percentual_sla_entrega': (df_entrega['DENTRO_SLA_ENTREGA'].sum() / len(df_entrega) * 100) if len(df_entrega) > 0 else 0
        }

        # Copiar outras análises do resultado original (simplificado)
        resultado_filtrado = {
            'df': df,
            'df_entrega': df_entrega,
            'stats_gerais': stats_gerais,
            'stats_entrega': stats_entrega,
            'analise_grupo': self.resultado_original.get('analise_grupo'),
            'analise_tipo': self.resultado_original.get('analise_tipo'),
            'analise_combinada': self.resultado_original.get('analise_combinada'),
            'analise_grupo_entrega': self.resultado_original.get('analise_grupo_entrega'),
            'analise_combinada_entrega': self.resultado_original.get('analise_combinada_entrega'),
            'distribuicao_tempo': self.resultado_original.get('distribuicao_tempo'),
            'analise_longitudinal_laudar': self.resultado_original.get('analise_longitudinal_laudar'),
            'analise_longitudinal_entrega': self.resultado_original.get('analise_longitudinal_entrega')
        }

        return resultado_filtrado

    def select_file(self):
        """Seleciona arquivo Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar Arquivo Excel",
            str(get_runtime_data_dir()),
            "Excel Files (*.xls *.xlsx)"
        )

        if file_path:
            self.current_file = file_path
            self.file_path_label.setText(Path(file_path).name)
            self.file_path_label.setStyleSheet("color: #58a6ff;")
            self.analyze_btn.setEnabled(True)

    def analyze_data(self):
        """Inicia análise de dados"""
        if not self.current_file:
            return

        start_date = self.start_date.date().toPython()
        end_date = self.end_date.date().toPython()

        # Desabilitar botão
        self.analyze_btn.setEnabled(False)
        self.analyze_btn.setText("⏳ Processando...")

        # Processar em thread separada
        self.processor = DataProcessor(self.current_file, start_date, end_date)
        self.processor.finished.connect(self.on_analysis_complete)
        self.processor.error.connect(self.on_analysis_error)
        self.processor.start()

    def on_analysis_complete(self, resultado):
        """Callback quando análise completa"""
        # Armazenar dados originais (sem filtro de hospital)
        self.resultado_original = resultado
        self.df_original = resultado['df']
        self.resultado = resultado
        self.df = resultado['df']

        # Extrair e popular dropdown de hospitais
        self.populate_hospital_dropdown()

        # Atualizar UI
        self.update_kpi_cards(resultado['stats_gerais'], resultado['stats_entrega'], resultado['df'])
        self.create_dashboard(resultado)

        # Habilitar botões
        self.search_patient_btn.setEnabled(True)
        self.generate_ai_btn.setEnabled(True)
        self.open_dashboard_btn.setEnabled(True)

        # Reabilitar botão
        self.analyze_btn.setEnabled(True)
        self.analyze_btn.setText("📊 Analisar Dados")

    def populate_hospital_dropdown(self):
        """Extrai hospitais únicos da coluna UNIDADE e popula dropdown"""
        if self.df_original is None:
            return

        # Verificar se coluna UNIDADE existe
        if 'UNIDADE' not in self.df_original.columns:
            # Se não existir, manter apenas "Todas as Unidades" e desabilitar
            self.hospital_combo.setEnabled(False)
            return

        # Extrair hospitais únicos (remover espaços extras e valores nulos)
        hospitais = self.df_original['UNIDADE'].dropna().str.strip().unique()
        hospitais = sorted([h for h in hospitais if h])  # Ordenar alfabeticamente

        # Limpar e repopular dropdown
        self.hospital_combo.blockSignals(True)  # Bloquear sinais temporariamente
        self.hospital_combo.clear()
        self.hospital_combo.addItem("Todas as Unidades")

        for hospital in hospitais:
            self.hospital_combo.addItem(hospital)

        # Restaurar seleção atual se existir
        index = self.hospital_combo.findText(self.selected_hospital)
        if index >= 0:
            self.hospital_combo.setCurrentIndex(index)
        else:
            self.hospital_combo.setCurrentIndex(0)
            self.selected_hospital = "Todas as Unidades"

        self.hospital_combo.blockSignals(False)  # Reabilitar sinais
        self.hospital_combo.setEnabled(True)  # Habilitar dropdown

    def on_analysis_error(self, error_msg):
        """Callback quando erro na análise"""
        QMessageBox.critical(self, "Erro", f"Erro ao processar dados:\n{error_msg}")
        self.analyze_btn.setEnabled(True)
        self.analyze_btn.setText("📊 Analisar Dados")

    def clear_layout(self, layout):
        """Remove widgets e layouts de um layout"""
        while layout.count():
            item = layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                self.clear_layout(item.layout())

    def update_kpi_cards(self, stats, stats_entrega, df):
        """Atualiza cards de KPI"""
        # Limpar cards existentes
        while self.kpi_layout.count():
            item = self.kpi_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # Linha 1: SLA de Atendimento (Laudar)
        row1 = QHBoxLayout()

        # Total de exames
        card1 = StatCard(
            "Total de Exames",
            f"{stats['total_exames']:,}",
            f"Período: {self.start_date.date().toString('dd/MM/yyyy')} - {self.end_date.date().toString('dd/MM/yyyy')}",
            "#2196F3"
        )

        # Tempo médio
        card2 = StatCard(
            "Tempo Médio Laudar",
            f"{stats['tempo_medio']:.1f} min",
            f"Mediana: {stats['tempo_mediano']:.1f} min",
            "#9C27B0"
        )

        # SLA Atendimento
        card3 = StatCard(
            "SLA de Realização (≤60min)",
            f"{stats['percentual_sla']:.1f}%",
            f"{stats['dentro_sla']:,} de {stats['total_exames']:,} exames",
            "#4CAF50" if stats['percentual_sla'] >= 80 else "#FF9800"
        )

        # Fora do SLA Atendimento
        card4 = StatCard(
            "Fora SLA de Realização",
            f"{stats['fora_sla']:,}",
            f"{100 - stats['percentual_sla']:.1f}% do total",
            "#F44336"
        )

        row1.addWidget(card1)
        row1.addWidget(card2)
        row1.addWidget(card3)
        row1.addWidget(card4)

        # Nova linha: Tempo Médio TC por Tipo de Atendimento
        row_tc = QHBoxLayout()

        # Calcular tempos médios de TC por tipo de atendimento
        df_tc = df[df['GRUPO'].str.upper().str.contains('TOMOGRAFIA', na=False)]

        # TC PA
        tc_pa = df_tc[df_tc['TIPO_ATENDIMENTO'].str.upper().str.contains('PRONTO', na=False)]
        tempo_tc_pa = tc_pa['TEMPO_ATENDIMENTO_MIN'].mean() if len(tc_pa) > 0 else 0
        card_tc_pa = StatCard(
            "TC - Pronto Atendimento",
            f"{tempo_tc_pa:.1f} min",
            f"{len(tc_pa):,} exames",
            "#FF6B6B"
        )

        # TC Externo
        tc_ext = df_tc[df_tc['TIPO_ATENDIMENTO'].str.upper().str.contains('EXTERNO', na=False)]
        tempo_tc_ext = tc_ext['TEMPO_ATENDIMENTO_MIN'].mean() if len(tc_ext) > 0 else 0
        card_tc_ext = StatCard(
            "TC - Externo",
            f"{tempo_tc_ext:.1f} min",
            f"{len(tc_ext):,} exames",
            "#51CF66"
        )

        # TC Internado
        tc_int = df_tc[df_tc['TIPO_ATENDIMENTO'].str.upper().str.contains('INTERNADO', na=False)]
        tempo_tc_int = tc_int['TEMPO_ATENDIMENTO_MIN'].mean() if len(tc_int) > 0 else 0
        card_tc_int = StatCard(
            "TC - Internado",
            f"{tempo_tc_int:.1f} min",
            f"{len(tc_int):,} exames",
            "#4E8FDF"
        )

        row_tc.addWidget(card_tc_pa)
        row_tc.addWidget(card_tc_ext)
        row_tc.addWidget(card_tc_int)

        # Nova linha: Tempo Médio RM por Tipo de Atendimento
        row_rm = QHBoxLayout()

        # Calcular tempos médios de RM por tipo de atendimento
        df_rm = df[df['GRUPO'].str.upper().str.contains('RESSONÂNCIA|RESSONANCIA|MAGNÉTICA', na=False, regex=True)]

        # RM PA
        rm_pa = df_rm[df_rm['TIPO_ATENDIMENTO'].str.upper().str.contains('PRONTO', na=False)]
        tempo_rm_pa = rm_pa['TEMPO_ATENDIMENTO_MIN'].mean() if len(rm_pa) > 0 else 0
        card_rm_pa = StatCard(
            "RM - Pronto Atendimento",
            f"{tempo_rm_pa:.1f} min",
            f"{len(rm_pa):,} exames",
            "#FF6B6B"
        )

        # RM Externo
        rm_ext = df_rm[df_rm['TIPO_ATENDIMENTO'].str.upper().str.contains('EXTERNO', na=False)]
        tempo_rm_ext = rm_ext['TEMPO_ATENDIMENTO_MIN'].mean() if len(rm_ext) > 0 else 0
        card_rm_ext = StatCard(
            "RM - Externo",
            f"{tempo_rm_ext:.1f} min",
            f"{len(rm_ext):,} exames",
            "#51CF66"
        )

        # RM Internado
        rm_int = df_rm[df_rm['TIPO_ATENDIMENTO'].str.upper().str.contains('INTERNADO', na=False)]
        tempo_rm_int = rm_int['TEMPO_ATENDIMENTO_MIN'].mean() if len(rm_int) > 0 else 0
        card_rm_int = StatCard(
            "RM - Internado",
            f"{tempo_rm_int:.1f} min",
            f"{len(rm_int):,} exames",
            "#4E8FDF"
        )

        row_rm.addWidget(card_rm_pa)
        row_rm.addWidget(card_rm_ext)
        row_rm.addWidget(card_rm_int)

        # Linha 2: SLA de Entrega de Resultado
        row2 = QHBoxLayout()

        # Total com resultado
        card5 = StatCard(
            "Exames com Resultado",
            f"{stats_entrega['total_exames']:,}",
            f"Com laudo preliminar ou aprovado",
            "#00BCD4"
        )

        # Tempo médio entrega
        card6 = StatCard(
            "Tempo Médio Entrega",
            f"{stats_entrega['tempo_medio_entrega']:.1f}h",
            f"≈ {stats_entrega['dias_uteis_medio']:.1f} dias úteis",
            "#673AB7"
        )

        # SLA Entrega
        card7 = StatCard(
            "SLA Entrega Resultado",
            f"{stats_entrega['percentual_sla_entrega']:.1f}%",
            f"{stats_entrega['dentro_sla_entrega']:,} dentro do prazo",
            "#4CAF50" if stats_entrega['percentual_sla_entrega'] >= 80 else "#FF9800"
        )

        # Fora SLA Entrega
        card8 = StatCard(
            "Fora SLA Entrega",
            f"{stats_entrega['fora_sla_entrega']:,}",
            f"{100 - stats_entrega['percentual_sla_entrega']:.1f}% do total",
            "#F44336"
        )

        row2.addWidget(card5)
        row2.addWidget(card6)
        row2.addWidget(card7)
        row2.addWidget(card8)

        # Adicionar separador TC/RM
        separator_tc_rm = QLabel("━━━ TEMPO MÉDIO POR MODALIDADE ━━━")
        separator_tc_rm.setAlignment(Qt.AlignCenter)
        separator_tc_rm.setStyleSheet("color: #58a6ff; font-weight: bold; font-size: 14px; padding: 10px;")

        # Adicionar separador
        separator = QLabel("━━━ SLA DE ENTREGA DE RESULTADO ━━━")
        separator.setAlignment(Qt.AlignCenter)
        separator.setStyleSheet("color: #58a6ff; font-weight: bold; font-size: 14px; padding: 10px;")

        # Adicionar ao layout principal
        self.kpi_layout.addLayout(row1)
        self.kpi_layout.addWidget(separator_tc_rm)
        self.kpi_layout.addLayout(row_tc)
        self.kpi_layout.addLayout(row_rm)
        self.kpi_layout.addWidget(separator)
        self.kpi_layout.addLayout(row2)

    def create_dashboard(self, resultado):
        """Cria dashboard com gráficos e tabelas"""
        # Limpar dashboard
        self.clear_layout(self.dashboard_layout)
        if hasattr(self, 'convenio_layout'):
            self.clear_layout(self.convenio_layout)

        # Primeira linha: Distribuição de tempo + SLA por modalidade
        row1 = QHBoxLayout()

        # Gráfico 1: Distribuição por faixa de tempo
        chart1 = self.create_time_distribution_chart(resultado['distribuicao_tempo'])
        row1.addWidget(chart1)

        # Gráfico 2: SLA por modalidade
        chart2 = self.create_sla_by_modality_chart(resultado['analise_grupo'])
        row1.addWidget(chart2)

        self.dashboard_layout.addLayout(row1)

        # Segunda linha: Tempo médio por modalidade + Tempo médio por tipo atendimento
        row2 = QHBoxLayout()

        chart3 = self.create_avg_time_by_modality_chart(resultado['analise_grupo'])
        row2.addWidget(chart3)

        chart4 = self.create_avg_time_by_type_chart(resultado['analise_tipo'])
        row2.addWidget(chart4)

        self.dashboard_layout.addLayout(row2)


        # Terceira linha: Tabela detalhada por modalidade
        table_group = QGroupBox("Análise Detalhada por Modalidade (GRUPO)")
        table_layout = QVBoxLayout()
        table1 = self.create_modality_table(resultado['analise_grupo'])
        table_layout.addWidget(table1)
        table_group.setLayout(table_layout)
        self.dashboard_layout.addWidget(table_group)

        # Quarta linha: Tabela por tipo de atendimento
        table_group2 = QGroupBox("Análise Detalhada por Tipo de Atendimento")
        table_layout2 = QVBoxLayout()
        table2 = self.create_type_table(resultado['analise_tipo'])
        table_layout2.addWidget(table2)
        table_group2.setLayout(table_layout2)
        self.dashboard_layout.addWidget(table_group2)

        # Nova tabela: Análise combinada (Grupo × Tipo)
        table_group3 = QGroupBox("Análise Combinada: Modalidade × Tipo de Atendimento")
        table_layout3 = QVBoxLayout()
        table3 = self.create_combined_table(resultado['analise_combinada'])
        table_layout3.addWidget(table3)
        table_group3.setLayout(table_layout3)
        self.dashboard_layout.addWidget(table_group3)

        # Separador: SLA DE ENTREGA DE RESULTADO
        separator = QLabel("═══════════ 📦 SLA DE ENTREGA DE RESULTADO ═══════════")
        separator.setAlignment(Qt.AlignCenter)
        separator.setStyleSheet("""
            color: #58a6ff;
            font-weight: bold;
            font-size: 18px;
            padding: 20px;
            background-color: #161b22;
            border-radius: 8px;
            margin: 20px 0px;
        """)
        self.dashboard_layout.addWidget(separator)

        # Primeira linha: Distribuição SLA + Distribuição por Faixa de Tempo
        entrega_row1 = QHBoxLayout()
        chart_entrega_dist = self.create_entrega_distribution_chart(resultado['df_entrega'])
        chart_entrega_time_dist = self.create_entrega_time_distribution_chart(resultado['df_entrega'])
        entrega_row1.addWidget(chart_entrega_dist)
        entrega_row1.addWidget(chart_entrega_time_dist)
        self.dashboard_layout.addLayout(entrega_row1)

        # Segunda linha: SLA por Modalidade + SLA por Tipo de Atendimento
        entrega_row2 = QHBoxLayout()
        chart_entrega_sla_mod = self.create_sla_entrega_by_modality_chart(resultado['analise_grupo_entrega'])
        chart_entrega_sla_type = self.create_entrega_sla_by_type_chart(resultado['analise_combinada_entrega'])
        entrega_row2.addWidget(chart_entrega_sla_mod)
        entrega_row2.addWidget(chart_entrega_sla_type)
        self.dashboard_layout.addLayout(entrega_row2)

        # Terceira linha: Tempo Médio por Modalidade + Tempo Médio por Tipo
        entrega_row3 = QHBoxLayout()
        chart_entrega_avg_mod = self.create_avg_entrega_time_by_modality_chart(resultado['analise_grupo_entrega'])
        chart_entrega_avg_type = self.create_avg_entrega_time_by_type_chart(resultado['analise_combinada_entrega'])
        entrega_row3.addWidget(chart_entrega_avg_mod)
        entrega_row3.addWidget(chart_entrega_avg_type)
        self.dashboard_layout.addLayout(entrega_row3)

        # Quarta linha: Heatmap combinado
        chart_entrega_heatmap = self.create_combined_entrega_chart(resultado['analise_combinada_entrega'])
        self.dashboard_layout.addWidget(chart_entrega_heatmap)

        # Tabela detalhada de SLA de Entrega
        table_group_entrega = QGroupBox("Análise Detalhada: SLA de Entrega de Resultado")
        table_layout_entrega = QVBoxLayout()
        table_entrega = self.create_entrega_table(resultado['analise_combinada_entrega'])
        table_layout_entrega.addWidget(table_entrega)
        table_group_entrega.setLayout(table_layout_entrega)
        self.dashboard_layout.addWidget(table_group_entrega)

        # Lista de pacientes fora do SLA de Entrega
        patient_list_entrega_group = QGroupBox("📋 Pacientes com Entrega Fora do Prazo")
        patient_list_entrega_layout = QVBoxLayout()

        export_entrega_btn = QPushButton("💾 Exportar Lista para Excel")
        export_entrega_btn.setObjectName("primaryButton")
        export_entrega_btn.clicked.connect(lambda: self.export_patient_entrega_list(resultado['df_entrega']))
        patient_list_entrega_layout.addWidget(export_entrega_btn)

        patient_entrega_table = self.create_patient_entrega_list_table(resultado['df_entrega'])
        patient_list_entrega_layout.addWidget(patient_entrega_table)
        patient_list_entrega_group.setLayout(patient_list_entrega_layout)
        self.dashboard_layout.addWidget(patient_list_entrega_group)

        # Separador
        separator2 = QLabel("═══════════ ⚠️ SLA DE ATENDIMENTO (LAUDAR) ═══════════")
        separator2.setAlignment(Qt.AlignCenter)
        separator2.setStyleSheet("""
            color: #58a6ff;
            font-weight: bold;
            font-size: 18px;
            padding: 20px;
            background-color: #161b22;
            border-radius: 8px;
            margin: 20px 0px;
        """)
        self.dashboard_layout.addWidget(separator2)

        # Nova tabela: Lista detalhada de pacientes fora do SLA de Realização
        patient_list_group = QGroupBox("📋 Pacientes Fora do Prazo Laudar (>60min)")
        patient_list_layout = QVBoxLayout()

        # Botão para exportar
        export_btn = QPushButton("💾 Exportar Lista para Excel")
        export_btn.setObjectName("primaryButton")
        export_btn.clicked.connect(lambda: self.export_patient_list(resultado['df']))
        patient_list_layout.addWidget(export_btn)

        patient_table = self.create_patient_list_table(resultado['df'])
        patient_list_layout.addWidget(patient_table)
        patient_list_group.setLayout(patient_list_layout)
        self.dashboard_layout.addWidget(patient_list_group)

        # Tab Convênios
        if hasattr(self, 'convenio_layout'):
            for item in self.create_convenio_section(resultado['df']):
                if isinstance(item, QLayout):
                    self.convenio_layout.addLayout(item)
                else:
                    self.convenio_layout.addWidget(item)

    def create_convenio_section(self, df):
        """Cria widgets de análise de convênios para a janela principal"""
        widgets = []

        separator = QLabel("═══════════ 🏥 ANÁLISE DE CONVÊNIOS ═══════════")
        separator.setAlignment(Qt.AlignCenter)
        separator.setStyleSheet("""
            color: #58a6ff;
            font-weight: bold;
            font-size: 18px;
            padding: 20px;
            background-color: #161b22;
            border-radius: 8px;
            margin: 20px 0px;
        """)
        widgets.append(separator)

        if df is None or len(df) == 0:
            widgets.append(QLabel("⚠️ Nenhum dado disponível para análise de convênios."))
            return widgets

        convenio_col = None
        for col in ['CONVENIO', 'CONVÊNIO', 'convenio', 'Convênio']:
            if col in df.columns:
                convenio_col = col
                break

        if convenio_col is None:
            widgets.append(QLabel("⚠️ Coluna CONVENIO não encontrada na planilha."))
            return widgets

        convenio_series = df[convenio_col].astype(str).str.strip()
        convenio_series = convenio_series.replace({'nan': '', 'NaT': '', 'None': '', 'NONE': ''})
        df_convenio = df[convenio_series != ''].copy()

        if len(df_convenio) == 0:
            widgets.append(QLabel("⚠️ Nenhum convênio preenchido encontrado."))
            return widgets

        df_convenio['CONVENIO_NORM'] = convenio_series[convenio_series != '']

        convenios_count = df_convenio['CONVENIO_NORM'].value_counts()

        # Linha 1: Top convênios + distribuição
        row1 = QHBoxLayout()

        # Gráfico de barras: Top convênios
        top_bar = convenios_count.head(12)[::-1]
        chart_bar = MatplotlibChart()
        ax = chart_bar.figure.add_subplot(111)

        labels = []
        for nome in top_bar.index:
            label = str(nome)
            if len(label) > 24:
                label = f"{label[:21]}..."
            labels.append(label)

        bars = ax.barh(range(len(top_bar)), top_bar.values,
                       color='#58a6ff', edgecolor='white', linewidth=1.5)
        ax.set_yticks(range(len(top_bar)))
        ax.set_yticklabels(labels, color='#c9d1d9')
        ax.set_xlabel('Quantidade de Exames', color='#c9d1d9', fontweight='bold')
        ax.set_title('Top Convênios por Volume', color='#58a6ff', fontweight='bold', pad=20)
        ax.tick_params(colors='#c9d1d9')
        ax.set_facecolor('#0d1117')
        ax.grid(True, alpha=0.2, axis='x', color='#30363d')

        for spine in ax.spines.values():
            spine.set_color('#30363d')

        for bar in bars:
            width = bar.get_width()
            ax.text(width + max(top_bar.values) * 0.01, bar.get_y() + bar.get_height()/2.,
                    f'{int(width):,}', ha='left', va='center',
                    color='#c9d1d9', fontweight='bold', fontsize=9)

        chart_bar.figure.tight_layout(pad=2.0)
        chart_bar.canvas.draw()

        container_bar = QGroupBox("📊 Convênios Mais Atendidos")
        layout_bar = QVBoxLayout()
        layout_bar.addWidget(chart_bar)
        container_bar.setLayout(layout_bar)
        row1.addWidget(container_bar)

        # Gráfico de pizza: distribuição
        chart_pie = MatplotlibChart()
        ax2 = chart_pie.figure.add_subplot(111)

        top_pie_n = min(8, len(convenios_count))
        pie_df = convenios_count.head(top_pie_n).reset_index()
        pie_df.columns = ['CONVENIO', 'Total']
        if len(convenios_count) > top_pie_n:
            outros_total = convenios_count.iloc[top_pie_n:].sum()
            pie_df = pd.concat(
                [pie_df, pd.DataFrame([{'CONVENIO': 'Outros', 'Total': outros_total}])],
                ignore_index=True
            )

        pie_labels = []
        for nome in pie_df['CONVENIO']:
            label = str(nome)
            if len(label) > 18:
                label = f"{label[:15]}..."
            pie_labels.append(label)

        colors_pie = ['#6f42c1', '#2196F3', '#00BCD4', '#4CAF50',
                      '#FF9800', '#E91E63', '#9C27B0', '#00A86B', '#FFC107']

        wedges, texts, autotexts = ax2.pie(
            pie_df['Total'].values,
            labels=pie_labels,
            autopct='%1.1f%%',
            colors=colors_pie[:len(pie_df)],
            textprops={'color': '#c9d1d9', 'fontsize': 9}
        )

        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(8)

        ax2.set_title('Distribuição de Convênios', color='#6f42c1', fontweight='bold', pad=20)
        chart_pie.figure.tight_layout(pad=2.0)
        chart_pie.canvas.draw()

        container_pie = QGroupBox("🥧 Distribuição Percentual")
        layout_pie = QVBoxLayout()
        layout_pie.addWidget(chart_pie)
        container_pie.setLayout(layout_pie)
        row1.addWidget(container_pie)

        widgets.append(row1)

        # Gráfico: Convênios por modalidade (Top 6 + Outros)
        if 'GRUPO' in df_convenio.columns:
            top_conv = convenios_count.head(6).index.tolist()
            df_stack = df_convenio.copy()
            df_stack['CONVENIO_TOP'] = np.where(
                df_stack['CONVENIO_NORM'].isin(top_conv),
                df_stack['CONVENIO_NORM'],
                'Outros'
            )

            pivot = df_stack.groupby(['GRUPO', 'CONVENIO_TOP']).size().unstack(fill_value=0)
            pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).index]

            cols_order = [c for c in top_conv if c in pivot.columns]
            if 'Outros' in pivot.columns:
                cols_order.append('Outros')
            pivot = pivot[cols_order]

            chart_stack = MatplotlibChart()
            ax3 = chart_stack.figure.add_subplot(111)

            num_modalidades = len(pivot)
            fig_height = max(6.0, num_modalidades * 0.5)
            chart_stack.figure.set_size_inches(10, fig_height, forward=True)

            y = range(num_modalidades)
            left = np.zeros(num_modalidades)
            colors_stack = ['#6f42c1', '#2196F3', '#00BCD4',
                            '#4CAF50', '#FF9800', '#E91E63', '#9C27B0']

            for i, convenio in enumerate(pivot.columns):
                values = pivot[convenio].values
                label = str(convenio)
                if len(label) > 16:
                    label = f"{label[:13]}..."
                ax3.barh(y, values, left=left,
                         label=label,
                         color=colors_stack[i % len(colors_stack)],
                         edgecolor='white', linewidth=0.5)
                left += values

            y_labels = []
            for modalidade in pivot.index:
                label = str(modalidade)
                if len(label) > 30:
                    label = f"{label[:27]}..."
                y_labels.append(label)

            ax3.set_yticks(y)
            ax3.set_yticklabels(y_labels, color='#c9d1d9', fontsize=9)
            ax3.set_xlabel('Quantidade de Exames', color='#c9d1d9', fontweight='bold')
            ax3.set_title('Convênios por Modalidade (Top 6 + Outros)',
                          color='#58a6ff', fontweight='bold', pad=20)
            ax3.tick_params(colors='#c9d1d9')
            ax3.set_facecolor('#0d1117')
            ax3.grid(True, alpha=0.2, axis='x', color='#30363d')
            ax3.legend(facecolor='#161b22', edgecolor='#30363d',
                       labelcolor='#c9d1d9', fontsize=8, loc='upper right')

            for spine in ax3.spines.values():
                spine.set_color('#30363d')

            chart_stack.figure.tight_layout(pad=2.0)
            chart_stack.canvas.draw()

            container_stack = QGroupBox("📌 Convênios x Modalidade")
            layout_stack = QVBoxLayout()
            layout_stack.addWidget(chart_stack)
            container_stack.setLayout(layout_stack)
            container_stack.setMinimumHeight(max(420, int(fig_height * 90)))
            widgets.append(container_stack)
        else:
            widgets.append(QLabel("⚠️ Coluna GRUPO não encontrada para análise por modalidade."))

        # Tabela: Top 3 convênios por modalidade
        if 'GRUPO' in df_convenio.columns:
            rows = []
            for modalidade, grupo_df in df_convenio.groupby('GRUPO'):
                counts_mod = grupo_df['CONVENIO_NORM'].value_counts()
                total_mod = counts_mod.sum()
                for convenio, total in counts_mod.head(3).items():
                    rows.append({
                        'Modalidade': modalidade,
                        'Convênio': convenio,
                        'Total': total,
                        'Percentual': (total / total_mod * 100) if total_mod else 0
                    })

            if rows:
                top_df = pd.DataFrame(rows).sort_values(['Modalidade', 'Total'], ascending=[True, False])

                table_group = QGroupBox("📋 Top Convênios por Modalidade")
                table_layout = QVBoxLayout()
                table = QTableWidget()
                table.setColumnCount(4)
                table.setHorizontalHeaderLabels(['Modalidade', 'Convênio', 'Total', 'Percentual'])
                table.setRowCount(len(top_df))
                table.setAlternatingRowColors(True)

                for i, (_, row) in enumerate(top_df.iterrows()):
                    table.setItem(i, 0, QTableWidgetItem(str(row['Modalidade'])))
                    table.setItem(i, 1, QTableWidgetItem(str(row['Convênio'])))
                    table.setItem(i, 2, QTableWidgetItem(f"{int(row['Total']):,}"))
                    table.setItem(i, 3, QTableWidgetItem(f"{row['Percentual']:.1f}%"))

                table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
                table.setMinimumHeight(260)
                table_layout.addWidget(table)
                table_group.setLayout(table_layout)
                widgets.append(table_group)

        return widgets

    def create_time_distribution_chart(self, distribuicao):
        """Gráfico de distribuição por faixa de tempo"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        colors = ['#4CAF50', '#8BC34A', '#CDDC39', '#FFC107', '#FF9800', '#FF5722', '#F44336']
        bars = ax.bar(range(len(distribuicao)), distribuicao.values, color=colors, edgecolor='white', linewidth=1.5)

        ax.set_xticks(range(len(distribuicao)))
        ax.set_xticklabels(distribuicao.index, rotation=45, ha='right', color='#c9d1d9')
        ax.set_ylabel('Quantidade de Exames', color='#c9d1d9', fontweight='bold')
        ax.set_title('Distribuição de Exames por Faixa de Tempo', color='#58a6ff', fontweight='bold', pad=20)
        ax.tick_params(colors='#c9d1d9')
        ax.spines['bottom'].set_color('#30363d')
        ax.spines['top'].set_color('#30363d')
        ax.spines['right'].set_color('#30363d')
        ax.spines['left'].set_color('#30363d')
        ax.set_facecolor('#0d1117')
        ax.grid(True, alpha=0.2, color='#30363d')

        # Adicionar valores nas barras
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{int(height)}',
                   ha='center', va='bottom', color='#c9d1d9', fontweight='bold')

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        container = QGroupBox("Distribuição por Faixa de Tempo")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    def create_sla_by_modality_chart(self, analise_grupo):
        """Gráfico de SLA por modalidade"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        # Calcular percentual SLA
        grupos = analise_grupo.index.tolist()
        total_exames = analise_grupo[('TEMPO_ATENDIMENTO_MIN', 'count')].values
        dentro_sla = analise_grupo[('DENTRO_SLA', 'sum')].values
        percentual_sla = (dentro_sla / total_exames * 100)

        # Ordenar por percentual
        sorted_indices = np.argsort(percentual_sla)[::-1]
        grupos = [grupos[i] for i in sorted_indices]
        percentual_sla = percentual_sla[sorted_indices]

        # Cores baseadas no percentual
        colors = ['#4CAF50' if p >= 80 else '#FF9800' if p >= 60 else '#F44336' for p in percentual_sla]

        bars = ax.barh(range(len(grupos)), percentual_sla, color=colors, edgecolor='white', linewidth=1.5)

        ax.set_yticks(range(len(grupos)))
        ax.set_yticklabels(grupos, color='#c9d1d9')
        ax.set_xlabel('% Dentro do SLA (≤60min)', color='#c9d1d9', fontweight='bold')
        ax.set_title('Taxa de Cumprimento do SLA por Modalidade', color='#58a6ff', fontweight='bold', pad=20)
        ax.tick_params(colors='#c9d1d9')
        ax.spines['bottom'].set_color('#30363d')
        ax.spines['top'].set_color('#30363d')
        ax.spines['right'].set_color('#30363d')
        ax.spines['left'].set_color('#30363d')
        ax.set_facecolor('#0d1117')
        ax.grid(True, alpha=0.2, axis='x', color='#30363d')
        ax.set_xlim(0, 100)

        # Adicionar valores
        for i, (bar, val) in enumerate(zip(bars, percentual_sla)):
            ax.text(val + 2, bar.get_y() + bar.get_height()/2.,
                   f'{val:.1f}%',
                   ha='left', va='center', color='#c9d1d9', fontweight='bold')

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        container = QGroupBox("Taxa de Cumprimento do SLA")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    def create_avg_time_by_modality_chart(self, analise_grupo):
        """Gráfico de tempo médio por modalidade"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        grupos = analise_grupo.index.tolist()
        tempo_medio = analise_grupo[('TEMPO_ATENDIMENTO_MIN', 'mean')].values

        # Ordenar por tempo
        sorted_indices = np.argsort(tempo_medio)[::-1]
        grupos = [grupos[i] for i in sorted_indices]
        tempo_medio = tempo_medio[sorted_indices]

        colors = ['#4CAF50' if t <= 60 else '#FF9800' if t <= 90 else '#F44336' for t in tempo_medio]

        bars = ax.barh(range(len(grupos)), tempo_medio, color=colors, edgecolor='white', linewidth=1.5)

        ax.set_yticks(range(len(grupos)))
        ax.set_yticklabels(grupos, color='#c9d1d9')
        ax.set_xlabel('Tempo Médio de Realização (minutos)', color='#c9d1d9', fontweight='bold')
        ax.set_title('Tempo Médio de Atendimento por Modalidade', color='#58a6ff', fontweight='bold', pad=20)
        ax.tick_params(colors='#c9d1d9')
        ax.spines['bottom'].set_color('#30363d')
        ax.spines['top'].set_color('#30363d')
        ax.spines['right'].set_color('#30363d')
        ax.spines['left'].set_color('#30363d')
        ax.set_facecolor('#0d1117')
        ax.grid(True, alpha=0.2, axis='x', color='#30363d')

        # Linha de referência SLA
        ax.axvline(x=60, color='#FFC107', linestyle='--', linewidth=2, label='SLA (60min)')
        ax.legend(facecolor='#161b22', edgecolor='#30363d', labelcolor='#c9d1d9')

        # Adicionar valores
        for i, (bar, val) in enumerate(zip(bars, tempo_medio)):
            ax.text(val + 2, bar.get_y() + bar.get_height()/2.,
                   f'{val:.1f}min',
                   ha='left', va='center', color='#c9d1d9', fontweight='bold')

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        container = QGroupBox("Tempo Médio por Modalidade")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    def create_avg_time_by_type_chart(self, analise_tipo):
        """Gráfico de tempo médio por tipo de atendimento"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        tipos = analise_tipo.index.tolist()
        tempo_medio = analise_tipo[('TEMPO_ATENDIMENTO_MIN', 'mean')].values

        colors = ['#2196F3', '#9C27B0', '#FF9800', '#4CAF50'][:len(tipos)]
        bars = ax.bar(range(len(tipos)), tempo_medio, color=colors, edgecolor='white', linewidth=1.5)

        ax.set_xticks(range(len(tipos)))
        ax.set_xticklabels(tipos, rotation=15, ha='right', color='#c9d1d9')
        ax.set_ylabel('Tempo Médio de Realização (minutos)', color='#c9d1d9', fontweight='bold')
        ax.set_title('Tempo Médio por Tipo de Atendimento', color='#58a6ff', fontweight='bold', pad=20)
        ax.tick_params(colors='#c9d1d9')
        ax.spines['bottom'].set_color('#30363d')
        ax.spines['top'].set_color('#30363d')
        ax.spines['right'].set_color('#30363d')
        ax.spines['left'].set_color('#30363d')
        ax.set_facecolor('#0d1117')
        ax.grid(True, alpha=0.2, color='#30363d')

        # Linha de referência SLA
        ax.axhline(y=60, color='#FFC107', linestyle='--', linewidth=2, label='SLA (60min)')
        ax.legend(facecolor='#161b22', edgecolor='#30363d', labelcolor='#c9d1d9')

        # Adicionar valores
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{height:.1f}min',
                   ha='center', va='bottom', color='#c9d1d9', fontweight='bold')

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        container = QGroupBox("Tempo Médio por Tipo de Atendimento")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    def create_modality_table(self, analise_grupo):
        """Tabela detalhada por modalidade"""
        table = QTableWidget()
        table.setColumnCount(7)
        table.setHorizontalHeaderLabels([
            'Modalidade', 'Qtd Exames', 'Tempo Médio', 'Tempo Mediano',
            'Tempo Min', 'Tempo Max', 'Dentro SLA'
        ])
        table.setRowCount(len(analise_grupo))
        table.setAlternatingRowColors(True)

        for i, (grupo, row) in enumerate(analise_grupo.iterrows()):
            table.setItem(i, 0, QTableWidgetItem(str(grupo)))
            table.setItem(i, 1, QTableWidgetItem(f"{int(row[('TEMPO_ATENDIMENTO_MIN', 'count')]):,}"))
            table.setItem(i, 2, QTableWidgetItem(f"{row[('TEMPO_ATENDIMENTO_MIN', 'mean')]:.1f} min"))
            table.setItem(i, 3, QTableWidgetItem(f"{row[('TEMPO_ATENDIMENTO_MIN', 'median')]:.1f} min"))
            table.setItem(i, 4, QTableWidgetItem(f"{row[('TEMPO_ATENDIMENTO_MIN', 'min')]:.1f} min"))
            table.setItem(i, 5, QTableWidgetItem(f"{row[('TEMPO_ATENDIMENTO_MIN', 'max')]:.1f} min"))

            dentro_sla = int(row[('DENTRO_SLA', 'sum')])
            total = int(row[('TEMPO_ATENDIMENTO_MIN', 'count')])
            perc = (dentro_sla / total * 100) if total > 0 else 0
            table.setItem(i, 6, QTableWidgetItem(f"{dentro_sla:,} ({perc:.1f}%)"))

        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setMinimumHeight(300)
        return table

    def create_type_table(self, analise_tipo):
        """Tabela detalhada por tipo de atendimento"""
        table = QTableWidget()
        table.setColumnCount(7)
        table.setHorizontalHeaderLabels([
            'Tipo Atendimento', 'Qtd Exames', 'Tempo Médio', 'Tempo Mediano',
            'Tempo Min', 'Tempo Max', 'Dentro SLA'
        ])
        table.setRowCount(len(analise_tipo))
        table.setAlternatingRowColors(True)

        for i, (tipo, row) in enumerate(analise_tipo.iterrows()):
            table.setItem(i, 0, QTableWidgetItem(str(tipo)))
            table.setItem(i, 1, QTableWidgetItem(f"{int(row[('TEMPO_ATENDIMENTO_MIN', 'count')]):,}"))
            table.setItem(i, 2, QTableWidgetItem(f"{row[('TEMPO_ATENDIMENTO_MIN', 'mean')]:.1f} min"))
            table.setItem(i, 3, QTableWidgetItem(f"{row[('TEMPO_ATENDIMENTO_MIN', 'median')]:.1f} min"))
            table.setItem(i, 4, QTableWidgetItem(f"{row[('TEMPO_ATENDIMENTO_MIN', 'min')]:.1f} min"))
            table.setItem(i, 5, QTableWidgetItem(f"{row[('TEMPO_ATENDIMENTO_MIN', 'max')]:.1f} min"))

            dentro_sla = int(row[('DENTRO_SLA', 'sum')])
            total = int(row[('TEMPO_ATENDIMENTO_MIN', 'count')])
            perc = (dentro_sla / total * 100) if total > 0 else 0
            table.setItem(i, 6, QTableWidgetItem(f"{dentro_sla:,} ({perc:.1f}%)"))

        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setMinimumHeight(200)
        return table

    def create_combined_heatmap(self, analise_combinada):
        """Mapa de calor com tempo médio por Grupo x Tipo de Atendimento"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        # Preparar dados para heatmap
        pivot_data = analise_combinada[('TEMPO_ATENDIMENTO_MIN', 'mean')].unstack(fill_value=0)

        # Criar heatmap
        im = ax.imshow(pivot_data.values, cmap='RdYlGn_r', aspect='auto', vmin=0, vmax=120)

        # Configurar eixos
        ax.set_xticks(np.arange(len(pivot_data.columns)))
        ax.set_yticks(np.arange(len(pivot_data.index)))
        ax.set_xticklabels(pivot_data.columns, rotation=45, ha='right', color='#c9d1d9')
        ax.set_yticklabels(pivot_data.index, color='#c9d1d9')

        # Adicionar valores nas células
        for i in range(len(pivot_data.index)):
            for j in range(len(pivot_data.columns)):
                value = pivot_data.values[i, j]
                if value > 0:
                    text_color = 'white' if value > 60 else 'black'
                    text = ax.text(j, i, f'{value:.1f}',
                                 ha="center", va="center", color=text_color, fontweight='bold')

        ax.set_title('Tempo Médio de Atendimento (minutos)', color='#58a6ff', fontweight='bold', pad=20)
        ax.set_xlabel('Tipo de Atendimento', color='#c9d1d9', fontweight='bold')
        ax.set_ylabel('Modalidade', color='#c9d1d9', fontweight='bold')

        # Colorbar
        cbar = chart_widget.figure.colorbar(im, ax=ax)
        cbar.set_label('Minutos', color='#c9d1d9', fontweight='bold')
        cbar.ax.tick_params(colors='#c9d1d9')

        ax.tick_params(colors='#c9d1d9')
        ax.set_facecolor('#0d1117')

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        return chart_widget

    def create_combined_table(self, analise_combinada):
        """Tabela combinada: Modalidade × Tipo de Atendimento"""
        table = QTableWidget()
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels([
            'Modalidade', 'Tipo Atendimento', 'Qtd Exames', 'Tempo Médio', '% Dentro SLA'
        ])

        # Resetar índice para poder iterar
        data = analise_combinada.reset_index()
        table.setRowCount(len(data))
        table.setAlternatingRowColors(True)

        for i, row in data.iterrows():
            grupo = row['GRUPO']
            tipo = row['TIPO_ATENDIMENTO']
            count = int(row[('TEMPO_ATENDIMENTO_MIN', 'count')])
            tempo_medio = row[('TEMPO_ATENDIMENTO_MIN', 'mean')]
            dentro_sla = int(row[('DENTRO_SLA', 'sum')])
            perc_sla = (dentro_sla / count * 100) if count > 0 else 0

            table.setItem(i, 0, QTableWidgetItem(str(grupo)))
            table.setItem(i, 1, QTableWidgetItem(str(tipo)))
            table.setItem(i, 2, QTableWidgetItem(f"{count:,}"))

            # Colorir tempo médio baseado no SLA
            tempo_item = QTableWidgetItem(f"{tempo_medio:.1f} min")
            if tempo_medio <= 60:
                tempo_item.setForeground(QColor("#4CAF50"))
            elif tempo_medio <= 90:
                tempo_item.setForeground(QColor("#FF9800"))
            else:
                tempo_item.setForeground(QColor("#F44336"))
            table.setItem(i, 3, tempo_item)

            # Colorir % SLA
            sla_item = QTableWidgetItem(f"{perc_sla:.1f}%")
            if perc_sla >= 80:
                sla_item.setForeground(QColor("#4CAF50"))
            elif perc_sla >= 60:
                sla_item.setForeground(QColor("#FF9800"))
            else:
                sla_item.setForeground(QColor("#F44336"))
            table.setItem(i, 4, sla_item)

        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setMinimumHeight(400)
        table.setSortingEnabled(True)
        return table

    def create_patient_list_table(self, df):
        """Tabela detalhada de pacientes fora do SLA"""
        # Filtrar apenas pacientes fora do SLA
        df_fora_sla = df[~df['DENTRO_SLA']].copy()
        df_fora_sla = df_fora_sla.sort_values('TEMPO_ATENDIMENTO_MIN', ascending=False)

        table = QTableWidget()
        table.setColumnCount(9)
        table.setHorizontalHeaderLabels([
            'SAME', 'Nome Paciente', 'Tipo Atendimento', 'Modalidade',
            'Data/Hora Prescrição', 'Data/Hora Laudo', 'Tempo (min)',
            'Atraso (min)', 'Status'
        ])

        table.setRowCount(len(df_fora_sla))
        table.setAlternatingRowColors(True)

        for i, (idx, row) in enumerate(df_fora_sla.iterrows()):
            # SAME
            same_item = QTableWidgetItem(str(int(row['SAME'])) if pd.notna(row['SAME']) else 'N/A')
            table.setItem(i, 0, same_item)

            # Nome
            nome_item = QTableWidgetItem(str(row['NOME_PACIENTE']) if pd.notna(row['NOME_PACIENTE']) else 'N/A')
            table.setItem(i, 1, nome_item)

            # Tipo Atendimento
            tipo_item = QTableWidgetItem(str(row['TIPO_ATENDIMENTO']) if pd.notna(row['TIPO_ATENDIMENTO']) else 'N/A')
            table.setItem(i, 2, tipo_item)

            # Modalidade
            grupo_item = QTableWidgetItem(str(row['GRUPO']) if pd.notna(row['GRUPO']) else 'N/A')
            table.setItem(i, 3, grupo_item)

            # Data/Hora Prescrição
            if pd.notna(row['DATA_HORA_PRESCRICAO']):
                prescricao_str = row['DATA_HORA_PRESCRICAO'].strftime('%d/%m/%Y %H:%M')
            else:
                prescricao_str = 'N/A'
            table.setItem(i, 4, QTableWidgetItem(prescricao_str))

            # Data/Hora Laudo
            if pd.notna(row['STATUS_ALAUDAR']):
                laudo_str = row['STATUS_ALAUDAR'].strftime('%d/%m/%Y %H:%M')
            else:
                laudo_str = 'N/A'
            table.setItem(i, 5, QTableWidgetItem(laudo_str))

            # Tempo total
            tempo_min = row['TEMPO_ATENDIMENTO_MIN']
            tempo_item = QTableWidgetItem(f"{tempo_min:.1f}")
            # Colorir baseado na gravidade do atraso
            if tempo_min > 120:
                tempo_item.setForeground(QColor("#F44336"))  # Vermelho
            elif tempo_min > 90:
                tempo_item.setForeground(QColor("#FF5722"))  # Laranja escuro
            else:
                tempo_item.setForeground(QColor("#FF9800"))  # Laranja
            table.setItem(i, 6, tempo_item)

            # Atraso (acima de 60min)
            atraso = tempo_min - 60
            atraso_item = QTableWidgetItem(f"{atraso:.1f}")
            atraso_item.setForeground(QColor("#F44336"))
            table.setItem(i, 7, atraso_item)

            # Status
            status_item = QTableWidgetItem("⚠️ FORA DO PRAZO")
            status_item.setForeground(QColor("#F44336"))
            table.setItem(i, 8, status_item)

        # Configurar colunas
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # SAME
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # Nome
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Tipo
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Modalidade
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Data Prescrição
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Data Laudo
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)  # Tempo
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)  # Atraso
        header.setSectionResizeMode(8, QHeaderView.ResizeToContents)  # Status

        table.setMinimumHeight(500)
        table.setSortingEnabled(True)

        # Label com total
        total_label = QLabel(f"Total de pacientes fora do prazo: {len(df_fora_sla)}")
        total_label.setStyleSheet("color: #F44336; font-weight: bold; font-size: 14px; padding: 5px;")

        # Container
        container = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(total_label)
        layout.addWidget(table)
        container.setLayout(layout)

        return container

    def export_patient_list(self, df):
        """Exporta lista de pacientes fora do SLA para Excel"""
        try:
            # Filtrar pacientes fora do SLA
            df_fora_sla = df[~df['DENTRO_SLA']].copy()

            if len(df_fora_sla) == 0:
                QMessageBox.information(self, "Exportar", "Não há pacientes fora do prazo para exportar.")
                return

            # Selecionar arquivo
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Salvar Lista de Pacientes",
                build_runtime_file_path(
                    f"pacientes_fora_prazo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                ),
                "Excel Files (*.xlsx)"
            )

            if not file_path:
                return

            # Preparar dados para exportação
            export_df = df_fora_sla[[
                'SAME', 'NOME_PACIENTE', 'TIPO_ATENDIMENTO', 'GRUPO',
                'DATA_HORA_PRESCRICAO', 'STATUS_ALAUDAR', 'TEMPO_ATENDIMENTO_MIN'
            ]].copy()

            export_df['ATRASO_MIN'] = export_df['TEMPO_ATENDIMENTO_MIN'] - 60
            export_df = export_df.sort_values('TEMPO_ATENDIMENTO_MIN', ascending=False)

            # Renomear colunas
            export_df.columns = [
                'SAME', 'Nome Paciente', 'Tipo Atendimento', 'Modalidade',
                'Data/Hora Prescrição', 'Data/Hora Laudo', 'Tempo Total (min)',
                'Atraso (min)'
            ]

            # Exportar
            export_df.to_excel(file_path, index=False, sheet_name='Pacientes Fora do Prazo')

            QMessageBox.information(
                self,
                "Exportação Concluída",
                f"Lista exportada com sucesso!\n\nArquivo: {file_path}\nTotal de registros: {len(df_fora_sla)}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Exportar", f"Erro ao exportar lista:\n{str(e)}")

    def create_sla_entrega_by_modality_chart(self, analise_grupo_entrega):
        """Gráfico de SLA de entrega por modalidade"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        grupos = analise_grupo_entrega.index.tolist()
        total_exames = analise_grupo_entrega[('TEMPO_ENTREGA_MIN', 'count')].values
        dentro_sla = analise_grupo_entrega[('DENTRO_SLA_ENTREGA', 'sum')].values
        percentual_sla = (dentro_sla / total_exames * 100)

        # Ordenar
        sorted_indices = np.argsort(percentual_sla)[::-1]
        grupos = [grupos[i] for i in sorted_indices]
        percentual_sla = percentual_sla[sorted_indices]

        colors = ['#4CAF50' if p >= 80 else '#FF9800' if p >= 60 else '#F44336' for p in percentual_sla]
        bars = ax.barh(range(len(grupos)), percentual_sla, color=colors, edgecolor='white', linewidth=1.5)

        ax.set_yticks(range(len(grupos)))
        ax.set_yticklabels(grupos, color='#c9d1d9')
        ax.set_xlabel('% Dentro do SLA de Entrega', color='#c9d1d9', fontweight='bold')
        ax.set_title('Taxa de Cumprimento SLA de Entrega por Modalidade', color='#58a6ff', fontweight='bold', pad=20)
        ax.tick_params(colors='#c9d1d9')
        ax.spines['bottom'].set_color('#30363d')
        ax.spines['top'].set_color('#30363d')
        ax.spines['right'].set_color('#30363d')
        ax.spines['left'].set_color('#30363d')
        ax.set_facecolor('#0d1117')
        ax.grid(True, alpha=0.2, axis='x', color='#30363d')
        ax.set_xlim(0, 100)

        for i, (bar, val) in enumerate(zip(bars, percentual_sla)):
            ax.text(val + 2, bar.get_y() + bar.get_height()/2.,
                   f'{val:.1f}%',
                   ha='left', va='center', color='#c9d1d9', fontweight='bold')

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        container = QGroupBox("SLA de Entrega por Modalidade")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    def create_combined_entrega_chart(self, analise_combinada_entrega):
        """Mapa de calor: Tempo médio de entrega"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        # Preparar dados
        pivot_data = analise_combinada_entrega[('TEMPO_ENTREGA_MIN', 'mean')].unstack(fill_value=0) / 60  # Converter para horas

        im = ax.imshow(pivot_data.values, cmap='RdYlGn_r', aspect='auto', vmin=0, vmax=168)  # 7 dias em horas

        ax.set_xticks(np.arange(len(pivot_data.columns)))
        ax.set_yticks(np.arange(len(pivot_data.index)))
        ax.set_xticklabels(pivot_data.columns, rotation=45, ha='right', color='#c9d1d9')
        ax.set_yticklabels(pivot_data.index, color='#c9d1d9')

        for i in range(len(pivot_data.index)):
            for j in range(len(pivot_data.columns)):
                value = pivot_data.values[i, j]
                if value > 0:
                    text_color = 'white' if value > 48 else 'black'
                    ax.text(j, i, f'{value:.1f}h',
                           ha="center", va="center", color=text_color, fontweight='bold', fontsize=9)

        ax.set_title('Tempo Médio de Entrega (horas)', color='#58a6ff', fontweight='bold', pad=20)
        ax.set_xlabel('Tipo de Atendimento', color='#c9d1d9', fontweight='bold')
        ax.set_ylabel('Modalidade', color='#c9d1d9', fontweight='bold')

        cbar = chart_widget.figure.colorbar(im, ax=ax)
        cbar.set_label('Horas', color='#c9d1d9', fontweight='bold')
        cbar.ax.tick_params(colors='#c9d1d9')

        ax.tick_params(colors='#c9d1d9')
        ax.set_facecolor('#0d1117')

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        container = QGroupBox("Tempo Médio de Entrega: Modalidade × Tipo")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    def create_entrega_distribution_chart(self, df_entrega):
        """Gráfico de pizza: Distribuição Dentro/Fora do SLA de Entrega"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        dentro_sla = df_entrega['DENTRO_SLA_ENTREGA'].sum()
        fora_sla = (~df_entrega['DENTRO_SLA_ENTREGA']).sum()

        sizes = [dentro_sla, fora_sla]
        labels = [f'Dentro do SLA\n{dentro_sla:,} exames', f'Fora do SLA\n{fora_sla:,} exames']
        colors = ['#4CAF50', '#F44336']
        explode = (0.05, 0.05)

        wedges, texts, autotexts = ax.pie(sizes, explode=explode, labels=labels, colors=colors,
                                           autopct='%1.1f%%', shadow=True, startangle=90,
                                           textprops={'color': '#c9d1d9', 'fontweight': 'bold', 'fontsize': 11})

        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(12)

        ax.set_title('Distribuição: Dentro vs Fora do SLA de Entrega', color='#58a6ff', fontweight='bold', pad=20)
        ax.set_facecolor('#0d1117')

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        container = QGroupBox("Distribuição SLA de Entrega")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    def create_avg_entrega_time_by_modality_chart(self, analise_grupo_entrega):
        """Gráfico de barras: Tempo médio de entrega por modalidade"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        grupos = analise_grupo_entrega.index.tolist()
        tempo_medio_h = analise_grupo_entrega[('TEMPO_ENTREGA_MIN', 'mean')].values / 60  # Converter para horas

        # Ordenar por tempo
        sorted_indices = np.argsort(tempo_medio_h)[::-1]
        grupos = [grupos[i] for i in sorted_indices]
        tempo_medio_h = tempo_medio_h[sorted_indices]

        colors = ['#4CAF50' if t <= 24 else '#FF9800' if t <= 72 else '#F44336' for t in tempo_medio_h]

        bars = ax.barh(range(len(grupos)), tempo_medio_h, color=colors, edgecolor='white', linewidth=1.5)

        ax.set_yticks(range(len(grupos)))
        ax.set_yticklabels(grupos, color='#c9d1d9')
        ax.set_xlabel('Tempo Médio de Entrega (horas)', color='#c9d1d9', fontweight='bold')
        ax.set_title('Tempo Médio de Entrega por Modalidade', color='#58a6ff', fontweight='bold', pad=20)
        ax.tick_params(colors='#c9d1d9')
        ax.spines['bottom'].set_color('#30363d')
        ax.spines['top'].set_color('#30363d')
        ax.spines['right'].set_color('#30363d')
        ax.spines['left'].set_color('#30363d')
        ax.set_facecolor('#0d1117')
        ax.grid(True, alpha=0.2, axis='x', color='#30363d')

        # Adicionar valores
        for i, (bar, val) in enumerate(zip(bars, tempo_medio_h)):
            ax.text(val + (max(tempo_medio_h) * 0.02), bar.get_y() + bar.get_height()/2.,
                   f'{val:.1f}h',
                   ha='left', va='center', color='#c9d1d9', fontweight='bold')

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        container = QGroupBox("Tempo Médio de Entrega por Modalidade")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    def create_avg_entrega_time_by_type_chart(self, analise_combinada_entrega):
        """Gráfico de barras: Tempo médio de entrega por tipo de atendimento"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        # Agrupar por tipo de atendimento
        tipo_data = analise_combinada_entrega.groupby(level=1).agg({
            ('TEMPO_ENTREGA_MIN', 'mean'): 'mean',
            ('TEMPO_ENTREGA_MIN', 'count'): 'sum'
        })

        tipos = tipo_data.index.tolist()
        tempo_medio_h = tipo_data[('TEMPO_ENTREGA_MIN', 'mean')].values / 60  # Converter para horas
        counts = tipo_data[('TEMPO_ENTREGA_MIN', 'count')].values

        # Ordenar por tempo
        sorted_indices = np.argsort(tempo_medio_h)[::-1]
        tipos = [tipos[i] for i in sorted_indices]
        tempo_medio_h = tempo_medio_h[sorted_indices]
        counts = counts[sorted_indices]

        colors = ['#4CAF50' if t <= 24 else '#FF9800' if t <= 72 else '#F44336' for t in tempo_medio_h]

        bars = ax.barh(range(len(tipos)), tempo_medio_h, color=colors, edgecolor='white', linewidth=1.5)

        ax.set_yticks(range(len(tipos)))
        ax.set_yticklabels(tipos, color='#c9d1d9')
        ax.set_xlabel('Tempo Médio de Entrega (horas)', color='#c9d1d9', fontweight='bold')
        ax.set_title('Tempo Médio de Entrega por Tipo de Atendimento', color='#58a6ff', fontweight='bold', pad=20)
        ax.tick_params(colors='#c9d1d9')
        ax.spines['bottom'].set_color('#30363d')
        ax.spines['top'].set_color('#30363d')
        ax.spines['right'].set_color('#30363d')
        ax.spines['left'].set_color('#30363d')
        ax.set_facecolor('#0d1117')
        ax.grid(True, alpha=0.2, axis='x', color='#30363d')

        # Adicionar valores e counts
        for i, (bar, val, count) in enumerate(zip(bars, tempo_medio_h, counts)):
            ax.text(val + (max(tempo_medio_h) * 0.02), bar.get_y() + bar.get_height()/2.,
                   f'{val:.1f}h (n={int(count)})',
                   ha='left', va='center', color='#c9d1d9', fontweight='bold', fontsize=9)

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        container = QGroupBox("Tempo Médio de Entrega por Tipo de Atendimento")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    def create_entrega_time_distribution_chart(self, df_entrega):
        """Gráfico de distribuição por faixa de tempo de entrega"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        # Criar faixas de tempo de entrega em horas
        df_temp = df_entrega.copy()
        df_temp['TEMPO_ENTREGA_HORAS'] = df_temp['TEMPO_ENTREGA_MIN'] / 60

        bins = [0, 12, 24, 48, 72, 120, 168, float('inf')]
        labels = ['0-12h', '12-24h', '24-48h', '48-72h', '72-120h', '120-168h', '>168h']
        df_temp['FAIXA_ENTREGA'] = pd.cut(df_temp['TEMPO_ENTREGA_HORAS'], bins=bins, labels=labels)
        distribuicao = df_temp['FAIXA_ENTREGA'].value_counts().sort_index()

        colors = ['#4CAF50', '#8BC34A', '#CDDC39', '#FFC107', '#FF9800', '#FF5722', '#F44336']
        bars = ax.bar(range(len(distribuicao)), distribuicao.values, color=colors, edgecolor='white', linewidth=1.5)

        ax.set_xticks(range(len(distribuicao)))
        ax.set_xticklabels(distribuicao.index, rotation=45, ha='right', color='#c9d1d9')
        ax.set_ylabel('Quantidade de Exames', color='#c9d1d9', fontweight='bold')
        ax.set_title('Distribuição de Exames por Faixa de Tempo de Entrega', color='#58a6ff', fontweight='bold', pad=20)
        ax.tick_params(colors='#c9d1d9')
        ax.spines['bottom'].set_color('#30363d')
        ax.spines['top'].set_color('#30363d')
        ax.spines['right'].set_color('#30363d')
        ax.spines['left'].set_color('#30363d')
        ax.set_facecolor('#0d1117')
        ax.grid(True, alpha=0.2, color='#30363d')

        # Adicionar valores nas barras
        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax.text(bar.get_x() + bar.get_width()/2., height,
                       f'{int(height)}',
                       ha='center', va='bottom', color='#c9d1d9', fontweight='bold')

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        container = QGroupBox("Distribuição por Faixa de Tempo de Entrega")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    def create_entrega_sla_by_type_chart(self, analise_combinada_entrega):
        """Gráfico de SLA de entrega por tipo de atendimento"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        # Agrupar por tipo de atendimento
        tipo_data = analise_combinada_entrega.groupby(level=1).agg({
            ('TEMPO_ENTREGA_MIN', 'count'): 'sum',
            ('DENTRO_SLA_ENTREGA', 'sum'): 'sum'
        })

        tipos = tipo_data.index.tolist()
        total_exames = tipo_data[('TEMPO_ENTREGA_MIN', 'count')].values
        dentro_sla = tipo_data[('DENTRO_SLA_ENTREGA', 'sum')].values
        percentual_sla = (dentro_sla / total_exames * 100)

        # Ordenar por percentual
        sorted_indices = np.argsort(percentual_sla)[::-1]
        tipos = [tipos[i] for i in sorted_indices]
        percentual_sla = percentual_sla[sorted_indices]

        colors = ['#4CAF50' if p >= 80 else '#FF9800' if p >= 60 else '#F44336' for p in percentual_sla]

        bars = ax.barh(range(len(tipos)), percentual_sla, color=colors, edgecolor='white', linewidth=1.5)

        ax.set_yticks(range(len(tipos)))
        ax.set_yticklabels(tipos, color='#c9d1d9')
        ax.set_xlabel('% Dentro do SLA de Entrega', color='#c9d1d9', fontweight='bold')
        ax.set_title('Taxa de Cumprimento SLA de Entrega por Tipo de Atendimento', color='#58a6ff', fontweight='bold', pad=20)
        ax.tick_params(colors='#c9d1d9')
        ax.spines['bottom'].set_color('#30363d')
        ax.spines['top'].set_color('#30363d')
        ax.spines['right'].set_color('#30363d')
        ax.spines['left'].set_color('#30363d')
        ax.set_facecolor('#0d1117')
        ax.grid(True, alpha=0.2, axis='x', color='#30363d')
        ax.set_xlim(0, 100)

        # Adicionar valores
        for i, (bar, val) in enumerate(zip(bars, percentual_sla)):
            ax.text(val + 2, bar.get_y() + bar.get_height()/2.,
                   f'{val:.1f}%',
                   ha='left', va='center', color='#c9d1d9', fontweight='bold')

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        container = QGroupBox("SLA de Entrega por Tipo de Atendimento")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    # ========== MÉTODOS DE ANÁLISE LONGITUDINAL MENSAL ==========

    def create_entrega_longitudinal_tc_chart(self, analise_longitudinal_entrega):
        """Gráfico de evolução mensal do SLA de Entrega - Tomografia"""
        return self._create_longitudinal_chart(analise_longitudinal_entrega, 'TC', 'Tomografia')

    def create_entrega_longitudinal_rm_chart(self, analise_longitudinal_entrega):
        """Gráfico de evolução mensal do SLA de Entrega - Ressonância Magnética"""
        return self._create_longitudinal_chart(analise_longitudinal_entrega, 'RM', 'Ressonância Magnética')

    def create_entrega_longitudinal_us_chart(self, analise_longitudinal_entrega):
        """Gráfico de evolução mensal do SLA de Entrega - Ultrassom"""
        return self._create_longitudinal_chart(analise_longitudinal_entrega, 'US', 'Ultrassom')

    def create_entrega_longitudinal_rx_chart(self, analise_longitudinal_entrega):
        """Gráfico de evolução mensal do SLA de Entrega - Raio-X"""
        return self._create_longitudinal_chart(analise_longitudinal_entrega, 'RX', 'Raio-X')

    def create_entrega_longitudinal_mn_chart(self, analise_longitudinal_entrega):
        """Gráfico de evolução mensal do SLA de Entrega - Medicina Nuclear"""
        return self._create_longitudinal_chart(analise_longitudinal_entrega, 'MN', 'Medicina Nuclear')

    def _create_longitudinal_chart(self, analise_longitudinal_entrega, modalidade_code, modalidade_nome):
        """Método auxiliar para criar gráfico longitudinal de uma modalidade específica"""
        chart_widget = MatplotlibChart()
        # Ajustar tamanho da figura para melhor visualização de timeline
        chart_widget.figure.set_figwidth(10)
        chart_widget.figure.set_figheight(5)
        ax = chart_widget.figure.add_subplot(111)

        # Filtrar dados da modalidade específica
        df_modalidade = analise_longitudinal_entrega[
            analise_longitudinal_entrega['MODALIDADE'] == modalidade_code
        ].copy()

        # Verificar se há dados
        if len(df_modalidade) == 0:
            ax.text(0.5, 0.5, f'Sem dados disponíveis para {modalidade_nome}',
                   ha='center', va='center', color='#c9d1d9', fontsize=14,
                   transform=ax.transAxes)
            ax.set_facecolor('#0d1117')
            ax.tick_params(colors='#c9d1d9')
            for spine in ax.spines.values():
                spine.set_color('#30363d')
            chart_widget.figure.tight_layout(pad=2.0)
            chart_widget.canvas.draw()
            container = QGroupBox(f"Evolução SLA Entrega - {modalidade_nome}")
            layout = QVBoxLayout()
            layout.addWidget(chart_widget)
            container.setLayout(layout)
            return container

        # Ordenar por data
        df_modalidade = df_modalidade.sort_values('ANO_MES')

        # Definir cores para cada tipo de atendimento
        color_map = {
            'PA': '#FF6B6B',           # Vermelho/laranja para PA (urgente)
            'INTERNADO': '#4E8FDF',    # Azul para internado
            'EXTERNO': '#51CF66'       # Verde para externo
        }

        # Função para mapear marcador baseado no tipo de atendimento
        def get_marker(tipo):
            tipo_upper = str(tipo).upper()
            if 'PRONTO' in tipo_upper or 'PA' in tipo_upper:
                return 'o'  # círculo
            elif 'INTERNADO' in tipo_upper:
                return 's'  # quadrado
            elif 'EXTERNO' in tipo_upper:
                return '^'  # triângulo
            elif 'AMBULAT' in tipo_upper:
                return 'D'  # diamante
            else:
                return 'o'  # círculo (default)

        # Função para mapear estilo de linha baseado no tipo de atendimento
        def get_linestyle(tipo):
            tipo_upper = str(tipo).upper()
            if 'PRONTO' in tipo_upper or 'PA' in tipo_upper:
                return '-'  # linha sólida
            elif 'INTERNADO' in tipo_upper:
                return '--'  # linha tracejada
            elif 'EXTERNO' in tipo_upper:
                return '-.'  # linha ponto-traço
            elif 'AMBULAT' in tipo_upper:
                return ':'  # linha pontilhada
            else:
                return '-'  # sólida (default)

        # Obter tipos de atendimento únicos presentes nos dados
        tipos_atendimento = df_modalidade['TIPO_ATENDIMENTO'].unique()

        # Obter lista ordenada de meses (usando ANO_MES para ordem cronológica)
        meses_unicos = df_modalidade[['ANO_MES', 'ANO_MES_STR']].drop_duplicates().sort_values('ANO_MES')
        all_meses = meses_unicos['ANO_MES_STR'].tolist()

        # Plotar linha para cada tipo de atendimento
        for tipo in tipos_atendimento:
            df_tipo = df_modalidade[df_modalidade['TIPO_ATENDIMENTO'] == tipo].copy()
            df_tipo = df_tipo.sort_values('ANO_MES')

            # Extrair dados
            meses = df_tipo['ANO_MES_STR'].tolist()
            percentuais = df_tipo['PERCENTUAL_SLA'].tolist()

            # Criar mapeamento de índices baseado na ordem cronológica
            indices = [all_meses.index(m) for m in meses]

            # Plotar linha com markers e estilos diferentes
            color = color_map.get(tipo, '#999999')
            marker = get_marker(tipo)
            linestyle = get_linestyle(tipo)

            ax.plot(indices, percentuais,
                   color=color, linewidth=3, marker=marker,
                   markersize=12, label=tipo, linestyle=linestyle,
                   markeredgecolor='white', markeredgewidth=2.5)

        # Linha de referência em 90%
        ax.axhline(y=90, color='#FFA500', linestyle='--', linewidth=2, alpha=0.7, label='Meta 90%')

        # Configurar eixos
        ax.set_xticks(range(len(all_meses)))
        ax.set_xticklabels(all_meses, rotation=45, ha='right', color='#c9d1d9')

        ax.set_xlabel('Período', color='#c9d1d9', fontweight='bold', fontsize=11)
        ax.set_ylabel('% Dentro do SLA', color='#c9d1d9', fontweight='bold', fontsize=11)
        ax.set_title(f'Evolução SLA Entrega - {modalidade_nome} por Tipo de Atendimento',
                    color='#58a6ff', fontweight='bold', pad=20, fontsize=13)

        # Configurar aparência
        ax.tick_params(colors='#c9d1d9')
        ax.set_facecolor('#0d1117')
        ax.spines['bottom'].set_color('#30363d')
        ax.spines['top'].set_color('#30363d')
        ax.spines['right'].set_color('#30363d')
        ax.spines['left'].set_color('#30363d')
        ax.grid(True, alpha=0.3, linestyle=':', color='#30363d')
        ax.set_ylim(0, 105)

        # Legenda
        legend = ax.legend(loc='best', framealpha=0.9, fancybox=True, shadow=True)
        legend.get_frame().set_facecolor('#161b22')
        legend.get_frame().set_edgecolor('#30363d')
        for text in legend.get_texts():
            text.set_color('#c9d1d9')

        chart_widget.figure.tight_layout(pad=2.0)
        chart_widget.canvas.draw()

        # Empacotar em container
        container = QGroupBox(f"📈 Evolução SLA Entrega - {modalidade_nome}")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    # ========== MÉTODOS DE SLA POR MODALIDADE E PORTA ==========

    def create_sla_by_modality_and_port_tc(self, analise_combinada_entrega):
        """Gráfico de SLA de Entrega: TC por Porta"""
        return self._create_sla_modality_port_chart(analise_combinada_entrega, 'TC', 'Tomografia')

    def create_sla_by_modality_and_port_rm(self, analise_combinada_entrega):
        """Gráfico de SLA de Entrega: RM por Porta"""
        return self._create_sla_modality_port_chart(analise_combinada_entrega, 'RM', 'Ressonância Magnética')

    def create_sla_by_modality_and_port_us(self, analise_combinada_entrega):
        """Gráfico de SLA de Entrega: US por Porta"""
        return self._create_sla_modality_port_chart(analise_combinada_entrega, 'US', 'Ultrassom')

    def create_sla_by_modality_and_port_rx(self, analise_combinada_entrega):
        """Gráfico de SLA de Entrega: RX por Porta"""
        return self._create_sla_modality_port_chart(analise_combinada_entrega, 'RX', 'Raio-X')

    def create_sla_by_modality_and_port_mama(self, analise_combinada_entrega):
        """Gráfico de SLA de Entrega: Mamografia por Porta"""
        return self._create_sla_modality_port_chart(analise_combinada_entrega, 'MAMOGRAFIA', 'Mamografia')

    def create_sla_by_modality_and_port_mn(self, analise_combinada_entrega):
        """Gráfico de SLA de Entrega: Medicina Nuclear por Porta"""
        return self._create_sla_modality_port_chart(analise_combinada_entrega, 'MEDICINA NUCLEAR', 'Medicina Nuclear')

    def _create_sla_modality_port_chart(self, analise_combinada_entrega, modalidade_key, modalidade_nome):
        """Método auxiliar para criar gráfico de SLA por porta para uma modalidade específica"""
        chart_widget = MatplotlibChart()
        ax = chart_widget.figure.add_subplot(111)

        # Filtrar dados da modalidade específica
        try:
            # Buscar grupos que contenham a modalidade (busca por substring)
            grupos_nivel0 = analise_combinada_entrega.index.get_level_values(0)

            # Mapear modalidade_key para padrões de busca
            padroes_busca = {
                'TC': ['TOMOGRAFIA'],
                'RM': ['RESSONÂNCIA', 'RESSONANCIA', 'MAGNÉTICA'],
                'US': ['ULTRASSOM', 'ULTRASOM', 'ECOGRAFIA'],
                'RX': ['RAIO'],
                'MAMOGRAFIA': ['MAMOGRAFIA'],
                'MEDICINA NUCLEAR': ['MEDICINA NUCLEAR', 'NUCLEAR', 'CINTILOGRAFIA']
            }

            # Buscar grupos que correspondam à modalidade
            grupos_encontrados = []
            if modalidade_key in padroes_busca:
                for grupo in grupos_nivel0.unique():
                    grupo_upper = str(grupo).upper()
                    for padrao in padroes_busca[modalidade_key]:
                        if padrao in grupo_upper:
                            grupos_encontrados.append(grupo)
                            break

            if not grupos_encontrados:
                # Modalidade não encontrada - criar gráfico vazio
                ax.text(0.5, 0.5, f'Sem dados disponíveis para {modalidade_nome}',
                       ha='center', va='center', color='#c9d1d9', fontsize=14,
                       transform=ax.transAxes)
                ax.set_facecolor('#0d1117')
                ax.tick_params(colors='#c9d1d9')
                for spine in ax.spines.values():
                    spine.set_color('#30363d')
                chart_widget.figure.tight_layout(pad=2.0)
                chart_widget.canvas.draw()
                container = QGroupBox(f"SLA Entrega: {modalidade_nome} por Porta")
                layout = QVBoxLayout()
                layout.addWidget(chart_widget)
                container.setLayout(layout)
                return container

            # Extrair dados de todos os grupos encontrados e agregar por porta
            # Criar dicionários para acumular dados por porta
            dados_por_porta = {}

            for grupo in grupos_encontrados:
                df_grupo = analise_combinada_entrega.loc[grupo]

                # Iterar sobre cada porta (tipo de atendimento) desse grupo
                for porta in df_grupo.index:
                    if porta not in dados_por_porta:
                        dados_por_porta[porta] = {
                            'total_exames': 0,
                            'dentro_sla': 0
                        }

                    # Acumular dados
                    dados_por_porta[porta]['total_exames'] += df_grupo.loc[porta][('TEMPO_ENTREGA_MIN', 'count')]
                    dados_por_porta[porta]['dentro_sla'] += df_grupo.loc[porta][('DENTRO_SLA_ENTREGA', 'sum')]

            # Converter para listas
            portas = list(dados_por_porta.keys())
            total_exames = [dados_por_porta[p]['total_exames'] for p in portas]
            dentro_sla = [dados_por_porta[p]['dentro_sla'] for p in portas]
            percentual_sla = [(d / t * 100) if t > 0 else 0 for d, t in zip(dentro_sla, total_exames)]

            # Converter para numpy arrays
            total_exames = np.array(total_exames)
            dentro_sla = np.array(dentro_sla)
            percentual_sla = np.array(percentual_sla)

            # Cores por porta
            color_map = {
                'PA': '#FF6B6B',
                'PRONTO ATENDIMENTO': '#FF6B6B',
                'INTERNADO': '#4E8FDF',
                'EXTERNO': '#51CF66',
                'AMBULATORIO': '#FFA500'
            }

            colors = []
            for porta in portas:
                porta_upper = str(porta).upper()
                if 'PRONTO' in porta_upper or 'PA' in porta_upper:
                    colors.append('#FF6B6B')
                elif 'INTERNADO' in porta_upper:
                    colors.append('#4E8FDF')
                elif 'EXTERNO' in porta_upper:
                    colors.append('#51CF66')
                elif 'AMBULAT' in porta_upper:
                    colors.append('#FFA500')
                else:
                    colors.append('#999999')

            # Criar gráfico de barras
            bars = ax.bar(range(len(portas)), percentual_sla, color=colors, edgecolor='white', linewidth=2)

            # Configurar eixos
            ax.set_xticks(range(len(portas)))
            ax.set_xticklabels(portas, rotation=15, ha='right', color='#c9d1d9', fontsize=10)
            ax.set_ylabel('% Dentro do SLA', color='#c9d1d9', fontweight='bold')
            ax.set_title(f'SLA de Entrega: {modalidade_nome} por Porta', color='#58a6ff', fontweight='bold', pad=20)
            ax.tick_params(colors='#c9d1d9')
            ax.spines['bottom'].set_color('#30363d')
            ax.spines['top'].set_color('#30363d')
            ax.spines['right'].set_color('#30363d')
            ax.spines['left'].set_color('#30363d')
            ax.set_facecolor('#0d1117')
            ax.grid(True, alpha=0.2, axis='y', color='#30363d')
            ax.set_ylim(0, 105)

            # Linha de referência (90%)
            ax.axhline(y=90, color='#FFA500', linestyle='--', linewidth=2, alpha=0.7, label='Meta 90%')

            # Adicionar valores e quantidade de exames nas barras
            for i, (bar, val, total) in enumerate(zip(bars, percentual_sla, total_exames)):
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2., height + 2,
                       f'{val:.1f}%\n({int(total)} exames)',
                       ha='center', va='bottom', color='#c9d1d9', fontweight='bold', fontsize=9)

            # Legenda
            legend = ax.legend(loc='upper right', framealpha=0.9, fancybox=True, shadow=True)
            legend.get_frame().set_facecolor('#161b22')
            legend.get_frame().set_edgecolor('#30363d')
            for text in legend.get_texts():
                text.set_color('#c9d1d9')

            chart_widget.figure.tight_layout(pad=2.0)
            chart_widget.canvas.draw()

        except Exception as e:
            # Erro ao processar - criar gráfico vazio com mensagem de erro
            ax.text(0.5, 0.5, f'Erro ao processar dados:\n{str(e)}',
                   ha='center', va='center', color='#FF6B6B', fontsize=12,
                   transform=ax.transAxes)
            ax.set_facecolor('#0d1117')
            ax.tick_params(colors='#c9d1d9')
            for spine in ax.spines.values():
                spine.set_color('#30363d')
            chart_widget.figure.tight_layout(pad=2.0)
            chart_widget.canvas.draw()

        # Empacotar em container
        container = QGroupBox(f"SLA Entrega: {modalidade_nome} por Porta")
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    # ========== MÉTODOS LONGITUDINAIS PARA LAUDAR (TEMPO DE REALIZAÇÃO) ==========

    def create_laudar_longitudinal_tc_chart(self, analise_longitudinal_laudar):
        """Gráfico de evolução diária do SLA de Realização - Tomografia"""
        return self._create_longitudinal_laudar_chart(analise_longitudinal_laudar, 'TC', 'Tomografia')

    def create_laudar_longitudinal_rm_chart(self, analise_longitudinal_laudar):
        """Gráfico de evolução diária do SLA de Realização - Ressonância Magnética"""
        return self._create_longitudinal_laudar_chart(analise_longitudinal_laudar, 'RM', 'Ressonância Magnética')

    def create_laudar_longitudinal_us_chart(self, analise_longitudinal_laudar):
        """Gráfico de evolução diária do SLA de Realização - Ultrassom"""
        return self._create_longitudinal_laudar_chart(analise_longitudinal_laudar, 'US', 'Ultrassom')

    def create_laudar_longitudinal_rx_chart(self, analise_longitudinal_laudar):
        """Gráfico de evolução diária do SLA de Realização - Raio-X"""
        return self._create_longitudinal_laudar_chart(analise_longitudinal_laudar, 'RX', 'Raio-X')

    def create_laudar_longitudinal_mn_chart(self, analise_longitudinal_laudar):
        """Gráfico de evolução diária do SLA de Realização - Medicina Nuclear"""
        return self._create_longitudinal_laudar_chart(analise_longitudinal_laudar, 'MN', 'Medicina Nuclear')

    def _create_longitudinal_laudar_chart(self, analise_longitudinal_laudar, modalidade_code, modalidade_nome):
        """Método auxiliar para criar gráfico longitudinal diário de laudar de uma modalidade específica"""
        chart_widget = MatplotlibChart()
        # Ajustar tamanho da figura para melhor visualização de timeline diária
        chart_widget.figure.set_figwidth(16)  # Aumentado de 12 para 16
        chart_widget.figure.set_figheight(7)  # Aumentado de 6 para 7
        ax1 = chart_widget.figure.add_subplot(111)

        # Criar segundo eixo Y para percentual SLA
        ax2 = ax1.twinx()

        # Filtrar dados da modalidade específica
        df_modalidade = analise_longitudinal_laudar[
            analise_longitudinal_laudar['MODALIDADE'] == modalidade_code
        ].copy()

        # Verificar se há dados
        if len(df_modalidade) == 0:
            ax1.text(0.5, 0.5, f'Sem dados disponíveis para {modalidade_nome}',
                   ha='center', va='center', color='#c9d1d9', fontsize=14,
                   transform=ax1.transAxes)
            ax1.set_facecolor('#0d1117')
            ax1.tick_params(colors='#c9d1d9')
            for spine in ax1.spines.values():
                spine.set_color('#30363d')
            ax2.tick_params(colors='#c9d1d9')
            for spine in ax2.spines.values():
                spine.set_color('#30363d')
            chart_widget.figure.tight_layout(pad=2.0)
            chart_widget.canvas.draw()
            container = QGroupBox(f"Evolução SLA de Realização - {modalidade_nome}")
            layout = QVBoxLayout()
            layout.addWidget(chart_widget)
            container.setLayout(layout)
            return container

        # Ordenar por data
        df_modalidade = df_modalidade.sort_values('DIA_PLANTAO')

        # Definir cores para cada tipo de atendimento
        color_map = {
            'PA': '#FF6B6B',           # Vermelho/laranja para PA (urgente)
            'INTERNADO': '#4E8FDF',    # Azul para internado
            'EXTERNO': '#51CF66'       # Verde para externo
        }

        # Função para mapear marcador baseado no tipo de atendimento
        def get_marker(tipo):
            tipo_upper = str(tipo).upper()
            if 'PRONTO' in tipo_upper or 'PA' in tipo_upper:
                return 'o'  # círculo
            elif 'INTERNADO' in tipo_upper:
                return 's'  # quadrado
            elif 'EXTERNO' in tipo_upper:
                return '^'  # triângulo
            elif 'AMBULAT' in tipo_upper:
                return 'D'  # diamante
            else:
                return 'o'  # círculo (default)

        # Função para mapear estilo de linha baseado no tipo de atendimento
        def get_linestyle(tipo):
            tipo_upper = str(tipo).upper()
            if 'PRONTO' in tipo_upper or 'PA' in tipo_upper:
                return '-'  # linha sólida
            elif 'INTERNADO' in tipo_upper:
                return '--'  # linha tracejada
            elif 'EXTERNO' in tipo_upper:
                return '-.'  # linha ponto-traço
            elif 'AMBULAT' in tipo_upper:
                return ':'  # linha pontilhada
            else:
                return '-'  # sólida (default)

        # Obter tipos de atendimento únicos presentes nos dados
        tipos_atendimento = df_modalidade['TIPO_ATENDIMENTO'].unique()

        # Plotar linhas para cada tipo de atendimento
        for tipo in tipos_atendimento:
            df_tipo = df_modalidade[df_modalidade['TIPO_ATENDIMENTO'] == tipo]

            # Extrair dados
            dias = df_tipo['DIA_STR'].tolist()
            tempos_medios = df_tipo['TEMPO_MEDIO'].tolist()
            percentuais = df_tipo['PERCENTUAL_SLA'].tolist()

            # Obter cor e estilos
            color = color_map.get(tipo, '#999999')
            marker = get_marker(tipo)
            linestyle_solid = get_linestyle(tipo)

            # Plotar TEMPO MÉDIO no eixo esquerdo (ax1)
            ax1.plot(range(len(dias)), tempos_medios,
                   color=color, linewidth=2, marker=marker,
                   markersize=6, label=f'{tipo}',
                   linestyle=linestyle_solid,
                   markeredgecolor='white', markeredgewidth=1)

        # Linha de referência em 60 minutos (SLA de 1 hora) no eixo esquerdo
        ax1.axhline(y=60, color='#FFA500', linestyle='--', linewidth=2, alpha=0.7, label='Meta 60 min')

        # Configurar eixos
        all_dias = sorted(df_modalidade['DIA_STR'].unique())

        # Se houver muitos dias (>15), mostrar apenas alguns labels para evitar sobreposição
        if len(all_dias) > 15:
            # Mostrar label a cada 2 ou 3 dias
            step = max(1, len(all_dias) // 10)  # Mostrar ~10 labels
            labels = [all_dias[i] if i % step == 0 else '' for i in range(len(all_dias))]
        else:
            labels = all_dias

        ax1.set_xticks(range(len(all_dias)))
        ax1.set_xticklabels(labels, rotation=90, ha='center', color='#c9d1d9', fontsize=8)

        ax1.set_xlabel('Dia (dd/mm)', color='#c9d1d9', fontweight='bold', fontsize=10)
        ax1.set_ylabel('Tempo Médio de Realização (minutos)', color='#c9d1d9', fontweight='bold', fontsize=10)
        ax1.set_title(f'Evolução Diária - Tempo Médio de Realização - {modalidade_nome}',
                    color='#58a6ff', fontweight='bold', pad=15, fontsize=12)

        # Configurar aparência
        ax1.tick_params(colors='#c9d1d9')
        ax1.set_facecolor('#0d1117')
        ax1.spines['bottom'].set_color('#30363d')
        ax1.spines['top'].set_color('#30363d')
        ax1.spines['right'].set_color('#30363d')
        ax1.spines['left'].set_color('#30363d')
        ax1.grid(True, alpha=0.3, linestyle=':', color='#30363d')

        # Ocultar eixo Y direito (ax2) já que não é mais usado
        ax2.set_visible(False)

        # Definir limites do eixo
        ax1.set_ylim(0, max(df_modalidade['TEMPO_MEDIO'].max() * 1.1, 70))

        # Legenda simplificada (apenas ax1)
        legend = ax1.legend(loc='upper left', framealpha=0.9, fancybox=True, shadow=True,
                          fontsize=9, ncol=1)  # 1 coluna agora
        legend.get_frame().set_facecolor('#161b22')
        legend.get_frame().set_edgecolor('#30363d')
        for text in legend.get_texts():
            text.set_color('#c9d1d9')

        # Ajustar layout com mais espaço para evitar crop de textos e legendas
        chart_widget.figure.subplots_adjust(left=0.10, right=0.95, top=0.92, bottom=0.18)
        chart_widget.canvas.draw()

        # Empacotar em container com altura mínima maior
        container = QGroupBox(f"📈 Evolução Diária - SLA de Realização - {modalidade_nome}")
        container.setMinimumHeight(650)  # Aumentar altura mínima
        layout = QVBoxLayout()
        layout.addWidget(chart_widget)
        container.setLayout(layout)
        return container

    # ========== FIM DOS MÉTODOS LONGITUDINAIS ==========

    def create_entrega_table(self, analise_combinada_entrega):
        """Tabela de SLA de entrega"""
        table = QTableWidget()
        table.setColumnCount(6)
        table.setHorizontalHeaderLabels([
            'Modalidade', 'Tipo Atendimento', 'Qtd Exames', 'Tempo Médio Entrega',
            'Dias Úteis Médio', '% Dentro SLA'
        ])

        data = analise_combinada_entrega.reset_index()
        table.setRowCount(len(data))
        table.setAlternatingRowColors(True)

        for i, row in data.iterrows():
            grupo = row['GRUPO']
            tipo = row['TIPO_ATENDIMENTO']
            count = int(row[('TEMPO_ENTREGA_MIN', 'count')])
            tempo_medio_h = row[('TEMPO_ENTREGA_MIN', 'mean')] / 60
            dias_uteis = row[('DIAS_UTEIS_ENTREGA', 'mean')]
            dentro_sla = int(row[('DENTRO_SLA_ENTREGA', 'sum')])
            perc_sla = (dentro_sla / count * 100) if count > 0 else 0

            table.setItem(i, 0, QTableWidgetItem(str(grupo)))
            table.setItem(i, 1, QTableWidgetItem(str(tipo)))
            table.setItem(i, 2, QTableWidgetItem(f"{count:,}"))

            tempo_item = QTableWidgetItem(f"{tempo_medio_h:.1f}h")
            if tempo_medio_h <= 24:
                tempo_item.setForeground(QColor("#4CAF50"))
            elif tempo_medio_h <= 72:
                tempo_item.setForeground(QColor("#FF9800"))
            else:
                tempo_item.setForeground(QColor("#F44336"))
            table.setItem(i, 3, tempo_item)

            table.setItem(i, 4, QTableWidgetItem(f"{dias_uteis:.1f}"))

            sla_item = QTableWidgetItem(f"{perc_sla:.1f}%")
            if perc_sla >= 80:
                sla_item.setForeground(QColor("#4CAF50"))
            elif perc_sla >= 60:
                sla_item.setForeground(QColor("#FF9800"))
            else:
                sla_item.setForeground(QColor("#F44336"))
            table.setItem(i, 5, sla_item)

        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setMinimumHeight(400)
        table.setSortingEnabled(True)
        return table

    def create_patient_entrega_list_table(self, df_entrega):
        """Tabela de pacientes fora do SLA de entrega"""
        df_fora_sla = df_entrega[~df_entrega['DENTRO_SLA_ENTREGA']].copy()
        df_fora_sla = df_fora_sla.sort_values('TEMPO_ENTREGA_MIN', ascending=False)

        table = QTableWidget()
        table.setColumnCount(12)
        table.setHorizontalHeaderLabels([
            'SAME', 'Nome Paciente', 'Tipo Atendimento', 'Modalidade', 'Descrição Procedimento',
            'Médico Laudador', 'Data Laudo', 'Data Entrega', 'Tempo (h)', 'Dias Úteis',
            'SLA Esperado', 'Status'
        ])

        table.setRowCount(len(df_fora_sla))
        table.setAlternatingRowColors(True)

        for i, (idx, row) in enumerate(df_fora_sla.iterrows()):
            table.setItem(i, 0, QTableWidgetItem(str(int(row['SAME'])) if pd.notna(row['SAME']) else 'N/A'))
            table.setItem(i, 1, QTableWidgetItem(str(row['NOME_PACIENTE']) if pd.notna(row['NOME_PACIENTE']) else 'N/A'))
            table.setItem(i, 2, QTableWidgetItem(str(row['TIPO_ATENDIMENTO']) if pd.notna(row['TIPO_ATENDIMENTO']) else 'N/A'))
            table.setItem(i, 3, QTableWidgetItem(str(row['GRUPO']) if pd.notna(row['GRUPO']) else 'N/A'))

            # Descrição Procedimento
            descricao = 'N/A'
            for col in ['DESCRICAO_PROCEDIMENTO', 'DESCRICAO PROCEDIMENTO', 'DS_PROCEDIMENTO', 'PROCEDIMENTO']:
                if col in row.index and pd.notna(row[col]):
                    descricao = str(row[col])
                    break
            table.setItem(i, 4, QTableWidgetItem(descricao))

            # Médico Laudador (prioriza MEDICO_LAUDO_DEFINITIVO, senão MEDICO_LAUDOO_PRELIMINAR)
            medico = 'N/A'
            if 'MEDICO_LAUDO_DEFINITIVO' in row.index and pd.notna(row['MEDICO_LAUDO_DEFINITIVO']):
                medico = str(row['MEDICO_LAUDO_DEFINITIVO'])
            elif 'MEDICO_LAUDOO_PRELIMINAR' in row.index and pd.notna(row['MEDICO_LAUDOO_PRELIMINAR']):
                medico = str(row['MEDICO_LAUDOO_PRELIMINAR'])
            table.setItem(i, 5, QTableWidgetItem(medico))

            if pd.notna(row['STATUS_ALAUDAR']):
                table.setItem(i, 6, QTableWidgetItem(row['STATUS_ALAUDAR'].strftime('%d/%m/%Y %H:%M')))
            else:
                table.setItem(i, 6, QTableWidgetItem('N/A'))

            if pd.notna(row['DATA_ENTREGA_RESULTADO']):
                table.setItem(i, 7, QTableWidgetItem(row['DATA_ENTREGA_RESULTADO'].strftime('%d/%m/%Y %H:%M')))
            else:
                table.setItem(i, 7, QTableWidgetItem('N/A'))

            tempo_h = row['TEMPO_ENTREGA_MIN'] / 60
            tempo_item = QTableWidgetItem(f"{tempo_h:.1f}")
            if tempo_h > 120:
                tempo_item.setForeground(QColor("#F44336"))
            elif tempo_h > 48:
                tempo_item.setForeground(QColor("#FF5722"))
            else:
                tempo_item.setForeground(QColor("#FF9800"))
            table.setItem(i, 8, tempo_item)

            dias_uteis = row['DIAS_UTEIS_ENTREGA']
            table.setItem(i, 9, QTableWidgetItem(f"{dias_uteis:.0f}" if pd.notna(dias_uteis) else 'N/A'))

            sla_desc = row['SLA_ESPERADO_DESC'] if 'SLA_ESPERADO_DESC' in row else 'N/A'
            table.setItem(i, 10, QTableWidgetItem(str(sla_desc)))

            status_item = QTableWidgetItem("⚠️ FORA DO PRAZO")
            status_item.setForeground(QColor("#F44336"))
            table.setItem(i, 11, status_item)

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        header.setSectionResizeMode(5, QHeaderView.Stretch)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(8, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(9, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(10, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(11, QHeaderView.ResizeToContents)

        table.setMinimumHeight(500)
        table.setSortingEnabled(True)

        total_label = QLabel(f"Total de pacientes com entrega fora do prazo: {len(df_fora_sla)}")
        total_label.setStyleSheet("color: #F44336; font-weight: bold; font-size: 14px; padding: 5px;")

        container = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(total_label)
        layout.addWidget(table)
        container.setLayout(layout)

        return container

    def export_patient_entrega_list(self, df_entrega):
        """Exporta lista de pacientes fora do SLA de entrega"""
        try:
            df_fora_sla = df_entrega[~df_entrega['DENTRO_SLA_ENTREGA']].copy()

            if len(df_fora_sla) == 0:
                QMessageBox.information(self, "Exportar", "Não há pacientes com entrega fora do prazo.")
                return

            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Salvar Lista de Pacientes - Entrega",
                build_runtime_file_path(
                    f"pacientes_entrega_fora_prazo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                ),
                "Excel Files (*.xlsx)"
            )

            if not file_path:
                return

            # Selecionar colunas base
            cols_base = ['SAME', 'NOME_PACIENTE', 'TIPO_ATENDIMENTO', 'GRUPO']

            # Adicionar coluna de descrição do procedimento
            descricao_col = None
            for col in ['DESCRICAO_PROCEDIMENTO', 'DESCRICAO PROCEDIMENTO', 'DS_PROCEDIMENTO', 'PROCEDIMENTO']:
                if col in df_fora_sla.columns:
                    cols_base.append(col)
                    descricao_col = col
                    break

            # Adicionar coluna do médico (prioriza MEDICO_LAUDO_DEFINITIVO)
            medico_col = None
            if 'MEDICO_LAUDO_DEFINITIVO' in df_fora_sla.columns:
                cols_base.append('MEDICO_LAUDO_DEFINITIVO')
                medico_col = 'MEDICO_LAUDO_DEFINITIVO'
            elif 'MEDICO_LAUDOO_PRELIMINAR' in df_fora_sla.columns:
                cols_base.append('MEDICO_LAUDOO_PRELIMINAR')
                medico_col = 'MEDICO_LAUDOO_PRELIMINAR'

            cols_base.extend(['STATUS_ALAUDAR', 'DATA_ENTREGA_RESULTADO',
                'TEMPO_ENTREGA_MIN', 'DIAS_UTEIS_ENTREGA', 'SLA_ESPERADO_DESC'])

            # Filtrar apenas colunas existentes
            cols_existentes = [c for c in cols_base if c in df_fora_sla.columns]
            export_df = df_fora_sla[cols_existentes].copy()

            export_df['TEMPO_ENTREGA_HORAS'] = export_df['TEMPO_ENTREGA_MIN'] / 60
            export_df = export_df.sort_values('TEMPO_ENTREGA_MIN', ascending=False)

            # Definir nomes das colunas dinamicamente
            col_names = ['SAME', 'Nome Paciente', 'Tipo Atendimento', 'Modalidade']
            if descricao_col:
                col_names.append('Descrição Procedimento')
            if medico_col:
                col_names.append('Médico Laudador')
            col_names.extend(['Data/Hora Laudo', 'Data/Hora Entrega',
                'Tempo Total (min)', 'Dias Úteis', 'SLA Esperado', 'Tempo Total (h)'])

            export_df.columns = col_names

            export_df.to_excel(file_path, index=False, sheet_name='Entrega Fora do Prazo')

            QMessageBox.information(
                self,
                "Exportação Concluída",
                f"Lista exportada com sucesso!\n\nArquivo: {file_path}\nTotal de registros: {len(df_fora_sla)}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Exportar", f"Erro ao exportar lista:\n{str(e)}")

    def create_rm_externos_table(self, df_rm_externos):
        """Tabela de RMs externas fora do SLA de realização"""
        df_sorted = df_rm_externos.sort_values('TEMPO_ATENDIMENTO_MIN', ascending=False)

        table = QTableWidget()
        table.setColumnCount(8)
        table.setHorizontalHeaderLabels([
            'SAME', 'Nome Paciente', 'Exame', 'Data Prescrição',
            'Data Laudo', 'Tempo (min)', 'Tempo (h)', 'Status'
        ])

        table.setRowCount(len(df_sorted))
        table.setAlternatingRowColors(True)

        for i, (idx, row) in enumerate(df_sorted.iterrows()):
            table.setItem(i, 0, QTableWidgetItem(str(int(row['SAME'])) if pd.notna(row['SAME']) else 'N/A'))
            table.setItem(i, 1, QTableWidgetItem(str(row['NOME_PACIENTE']) if pd.notna(row['NOME_PACIENTE']) else 'N/A'))
            table.setItem(i, 2, QTableWidgetItem(str(row['GRUPO']) if pd.notna(row['GRUPO']) else 'N/A'))

            if pd.notna(row['DATA_HORA_PRESCRICAO']):
                table.setItem(i, 3, QTableWidgetItem(row['DATA_HORA_PRESCRICAO'].strftime('%d/%m/%Y %H:%M')))
            else:
                table.setItem(i, 3, QTableWidgetItem('N/A'))

            if pd.notna(row['STATUS_ALAUDAR']):
                table.setItem(i, 4, QTableWidgetItem(row['STATUS_ALAUDAR'].strftime('%d/%m/%Y %H:%M')))
            else:
                table.setItem(i, 4, QTableWidgetItem('N/A'))

            tempo_min = row['TEMPO_ATENDIMENTO_MIN']
            tempo_min_item = QTableWidgetItem(f"{tempo_min:.0f}")
            if tempo_min > 180:
                tempo_min_item.setForeground(QColor("#F44336"))
            elif tempo_min > 120:
                tempo_min_item.setForeground(QColor("#FF5722"))
            else:
                tempo_min_item.setForeground(QColor("#FF9800"))
            table.setItem(i, 5, tempo_min_item)

            tempo_h = tempo_min / 60
            tempo_h_item = QTableWidgetItem(f"{tempo_h:.1f}")
            if tempo_h > 3:
                tempo_h_item.setForeground(QColor("#F44336"))
            elif tempo_h > 2:
                tempo_h_item.setForeground(QColor("#FF5722"))
            else:
                tempo_h_item.setForeground(QColor("#FF9800"))
            table.setItem(i, 6, tempo_h_item)

            status_item = QTableWidgetItem("🚨 FORA DO PRAZO (>60 min)")
            status_item.setForeground(QColor("#F44336"))
            table.setItem(i, 7, status_item)

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)

        table.setMinimumHeight(400)
        table.setSortingEnabled(True)

        total_label = QLabel(f"Total de RMs Externas fora do prazo: {len(df_sorted)}")
        total_label.setStyleSheet("color: #F44336; font-weight: bold; font-size: 14px; padding: 5px;")

        container = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(total_label)
        layout.addWidget(table)
        container.setLayout(layout)

        return container

    def export_rm_externos_list(self, df_rm_externos):
        """Exporta lista de RMs externas fora do SLA"""
        try:
            if len(df_rm_externos) == 0:
                QMessageBox.information(self, "Exportar", "Não há RMs externas fora do prazo.")
                return

            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Salvar Lista de RMs Externas",
                build_runtime_file_path(
                    f"rm_externos_fora_prazo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                ),
                "Excel Files (*.xlsx)"
            )

            if not file_path:
                return

            export_df = df_rm_externos[[
                'SAME', 'NOME_PACIENTE', 'GRUPO', 'TIPO_ATENDIMENTO',
                'DATA_HORA_PRESCRICAO', 'STATUS_ALAUDAR', 'TEMPO_ATENDIMENTO_MIN'
            ]].copy()

            export_df['TEMPO_ATENDIMENTO_HORAS'] = export_df['TEMPO_ATENDIMENTO_MIN'] / 60
            export_df = export_df.sort_values('TEMPO_ATENDIMENTO_MIN', ascending=False)

            export_df.columns = [
                'SAME', 'Nome Paciente', 'Exame', 'Tipo Atendimento',
                'Data/Hora Prescrição', 'Data/Hora Laudo',
                'Tempo Total (min)', 'Tempo Total (h)'
            ]

            export_df.to_excel(file_path, index=False, sheet_name='RM Externos Fora do Prazo')

            QMessageBox.information(
                self,
                "Exportação Concluída",
                f"Lista exportada com sucesso!\n\nArquivo: {file_path}\nTotal de registros: {len(df_rm_externos)}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Exportar", f"Erro ao exportar lista:\n{str(e)}")

    def create_rm_pacientes_table(self, pacientes_agrupados):
        """Tabela de pacientes com RMs externas fora do SLA (agrupado)"""
        table = QTableWidget()
        table.setColumnCount(7)
        table.setHorizontalHeaderLabels([
            'SAME', 'Nome Paciente', 'Qtd Exames', 'Tempo Médio (min)',
            'Tempo Máximo (min)', 'Primeira Prescrição', 'Último Laudo'
        ])

        table.setRowCount(len(pacientes_agrupados))
        table.setAlternatingRowColors(True)

        for i, (idx, row) in enumerate(pacientes_agrupados.iterrows()):
            table.setItem(i, 0, QTableWidgetItem(str(int(row['SAME'])) if pd.notna(row['SAME']) else 'N/A'))
            table.setItem(i, 1, QTableWidgetItem(str(row['NOME_PACIENTE']) if pd.notna(row['NOME_PACIENTE']) else 'N/A'))

            qtd_item = QTableWidgetItem(str(int(row['QTD_EXAMES'])))
            if row['QTD_EXAMES'] > 2:
                qtd_item.setForeground(QColor("#F44336"))
                qtd_item.setFont(QFont("Arial", 10, QFont.Bold))
            elif row['QTD_EXAMES'] > 1:
                qtd_item.setForeground(QColor("#FF9800"))
            table.setItem(i, 2, qtd_item)

            tempo_medio = row['TEMPO_MEDIO']
            tempo_medio_item = QTableWidgetItem(f"{tempo_medio:.0f}")
            if tempo_medio > 180:
                tempo_medio_item.setForeground(QColor("#F44336"))
            elif tempo_medio > 120:
                tempo_medio_item.setForeground(QColor("#FF5722"))
            else:
                tempo_medio_item.setForeground(QColor("#FF9800"))
            table.setItem(i, 3, tempo_medio_item)

            tempo_max = row['TEMPO_MAXIMO']
            tempo_max_item = QTableWidgetItem(f"{tempo_max:.0f}")
            if tempo_max > 240:
                tempo_max_item.setForeground(QColor("#F44336"))
            elif tempo_max > 180:
                tempo_max_item.setForeground(QColor("#FF5722"))
            else:
                tempo_max_item.setForeground(QColor("#FF9800"))
            table.setItem(i, 4, tempo_max_item)

            if pd.notna(row['PRIMEIRA_PRESCRICAO']):
                table.setItem(i, 5, QTableWidgetItem(row['PRIMEIRA_PRESCRICAO'].strftime('%d/%m/%Y %H:%M')))
            else:
                table.setItem(i, 5, QTableWidgetItem('N/A'))

            if pd.notna(row['ULTIMO_LAUDO']):
                table.setItem(i, 6, QTableWidgetItem(row['ULTIMO_LAUDO'].strftime('%d/%m/%Y %H:%M')))
            else:
                table.setItem(i, 6, QTableWidgetItem('N/A'))

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)

        table.setMinimumHeight(400)
        table.setSortingEnabled(True)

        total_label = QLabel(f"Total de pacientes: {len(pacientes_agrupados)}")
        total_label.setStyleSheet("color: #F44336; font-weight: bold; font-size: 14px; padding: 5px;")

        container = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(total_label)
        layout.addWidget(table)
        container.setLayout(layout)

        return container

    def export_rm_pacientes_list(self, pacientes_agrupados):
        """Exporta lista de pacientes com RMs externas em atraso"""
        try:
            if len(pacientes_agrupados) == 0:
                QMessageBox.information(self, "Exportar", "Não há pacientes com RM externa em atraso.")
                return

            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Salvar Lista de Pacientes",
                build_runtime_file_path(
                    f"pacientes_rm_externos_atraso_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                ),
                "Excel Files (*.xlsx)"
            )

            if not file_path:
                return

            export_df = pacientes_agrupados.copy()
            export_df['TEMPO_MEDIO_HORAS'] = export_df['TEMPO_MEDIO'] / 60
            export_df['TEMPO_MAXIMO_HORAS'] = export_df['TEMPO_MAXIMO'] / 60
            export_df = export_df.sort_values('QTD_EXAMES', ascending=False)

            export_df.columns = [
                'SAME', 'Nome Paciente', 'Qtd Exames em Atraso',
                'Tempo Médio (min)', 'Tempo Máximo (min)',
                'Primeira Prescrição', 'Último Laudo',
                'Tempo Médio (h)', 'Tempo Máximo (h)'
            ]

            export_df.to_excel(file_path, index=False, sheet_name='Pacientes RM Externos')

            QMessageBox.information(
                self,
                "Exportação Concluída",
                f"Lista exportada com sucesso!\n\nArquivo: {file_path}\nTotal de pacientes: {len(pacientes_agrupados)}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Exportar", f"Erro ao exportar lista:\n{str(e)}")

    def open_patient_search(self):
        """Abre janela de busca de paciente"""
        if self.resultado is None:
            QMessageBox.warning(self, "Aviso", "Por favor, carregue e analise os dados primeiro.")
            return

        self.patient_search_window = PatientSearchWindow(self.resultado['df'], self)
        self.patient_search_window.show()

    def open_us_analysis(self):
        """Abre janela de análise estratégica de ultrassonografia"""
        if self.df is None:
            QMessageBox.warning(self, "Aviso", "Por favor, carregue e analise os dados primeiro.")
            return

        # Obter período selecionado
        start_date = self.start_date.date().toPython()
        end_date = self.end_date.date().toPython()

        # Obter nome da unidade selecionada
        unidade_nome = self.selected_hospital

        self.us_analysis_window = UltrasoundAnalysisWindow(self.df, start_date, end_date, unidade_nome, self)

        # Verificar se há exames de US
        if self.us_analysis_window.df_us is None or len(self.us_analysis_window.df_us) == 0:
            QMessageBox.warning(self, "Aviso",
                              "Não foram encontrados exames de ULTRASSOM no período selecionado.\n\n"
                              "Verifique se a coluna GRUPO contém registros com 'ULTRASSOM'.")
            return

        self.us_analysis_window.show()

    def open_dashboard_window(self):
        """Abre janela separada com dashboards"""
        if not self.resultado:
            QMessageBox.warning(self, "Aviso", "Execute a análise de dados primeiro!")
            return

        # Criar janela de dashboard
        self.dashboard_window = DashboardWindow(self)

        # Criar widgets para SLA de Realização
        laudar_widgets = []

        # Gráficos de SLA de Realização
        row1 = QWidget()
        row1_layout = QHBoxLayout(row1)
        chart1 = self.create_time_distribution_chart(self.resultado['distribuicao_tempo'])
        chart2 = self.create_sla_by_modality_chart(self.resultado['analise_grupo'])
        row1_layout.addWidget(chart1)
        row1_layout.addWidget(chart2)
        laudar_widgets.append(row1)

        row2 = QWidget()
        row2_layout = QHBoxLayout(row2)
        chart3 = self.create_avg_time_by_modality_chart(self.resultado['analise_grupo'])
        chart4 = self.create_avg_time_by_type_chart(self.resultado['analise_tipo'])
        row2_layout.addWidget(chart3)
        row2_layout.addWidget(chart4)
        laudar_widgets.append(row2)

        # Tabelas de SLA de Realização
        table_group1 = QGroupBox("Análise Detalhada por Modalidade")
        table_layout1 = QVBoxLayout()
        table1 = self.create_modality_table(self.resultado['analise_grupo'])
        table_layout1.addWidget(table1)
        table_group1.setLayout(table_layout1)
        laudar_widgets.append(table_group1)

        table_group2 = QGroupBox("Análise Detalhada por Tipo de Atendimento")
        table_layout2 = QVBoxLayout()
        table2 = self.create_type_table(self.resultado['analise_tipo'])
        table_layout2.addWidget(table2)
        table_group2.setLayout(table_layout2)
        laudar_widgets.append(table_group2)

        table_group3 = QGroupBox("Análise Combinada: Modalidade × Tipo")
        table_layout3 = QVBoxLayout()
        table3 = self.create_combined_table(self.resultado['analise_combinada'])
        table_layout3.addWidget(table3)
        table_group3.setLayout(table_layout3)
        laudar_widgets.append(table_group3)

        # ========== SEÇÃO DE ANÁLISE LONGITUDINAL DIÁRIA - LAUDAR ==========
        # Título da seção
        longitudinal_laudar_title = QLabel("📈 Análise Longitudinal Diária - Tempo de Realização por Modalidade")
        longitudinal_laudar_title.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #58a6ff;
                padding: 20px 10px 10px 10px;
                background-color: #161b22;
                border-radius: 8px;
                margin-top: 20px;
            }
        """)
        laudar_widgets.append(longitudinal_laudar_title)

        # Um gráfico por linha para melhor visualização
        # Linha 1: TC
        chart_laudar_tc = self.create_laudar_longitudinal_tc_chart(self.resultado['analise_longitudinal_laudar'])
        laudar_widgets.append(chart_laudar_tc)

        # Linha 2: RM
        chart_laudar_rm = self.create_laudar_longitudinal_rm_chart(self.resultado['analise_longitudinal_laudar'])
        laudar_widgets.append(chart_laudar_rm)

        # Linha 3: US
        chart_laudar_us = self.create_laudar_longitudinal_us_chart(self.resultado['analise_longitudinal_laudar'])
        laudar_widgets.append(chart_laudar_us)

        # Linha 4: RX
        chart_laudar_rx = self.create_laudar_longitudinal_rx_chart(self.resultado['analise_longitudinal_laudar'])
        laudar_widgets.append(chart_laudar_rx)

        # Linha 5: MN
        chart_laudar_mn = self.create_laudar_longitudinal_mn_chart(self.resultado['analise_longitudinal_laudar'])
        laudar_widgets.append(chart_laudar_mn)

        # ========== FIM DA SEÇÃO LONGITUDINAL LAUDAR ==========

        # Seção específica: RM Externos Fora do Prazo
        rm_externos_group = QGroupBox("🚨 Ressonâncias Magnéticas - Pacientes Externos Fora do Prazo de Realização")
        rm_externos_layout = QVBoxLayout()

        # Filtrar dados
        df_rm_externos = self.resultado['df'][
            (self.resultado['df']['GRUPO'].str.upper().str.contains('RESSONÂNCIA|RESSONANCIA|MAGNÉTICA', na=False, regex=True)) &
            (self.resultado['df']['TIPO_ATENDIMENTO'].str.upper().str.contains('EXTERNO', na=False)) &
            (self.resultado['df']['DENTRO_SLA'] == False)
        ]

        # Card indicador
        rm_externos_card_layout = QHBoxLayout()
        rm_externos_count = len(df_rm_externos)
        rm_externos_total = len(self.resultado['df'][
            (self.resultado['df']['GRUPO'].str.upper().str.contains('RESSONÂNCIA|RESSONANCIA|MAGNÉTICA', na=False, regex=True)) &
            (self.resultado['df']['TIPO_ATENDIMENTO'].str.upper().str.contains('EXTERNO', na=False))
        ])
        rm_externos_perc = (rm_externos_count / rm_externos_total * 100) if rm_externos_total > 0 else 0

        card_rm_externos = StatCard(
            "RM Externos Fora do Prazo",
            f"{rm_externos_count}",
            f"{rm_externos_perc:.1f}% do total de RM externos",
            "#FF6B6B"
        )
        rm_externos_card_layout.addWidget(card_rm_externos)
        rm_externos_card_layout.addStretch()

        card_widget = QWidget()
        card_widget.setLayout(rm_externos_card_layout)
        rm_externos_layout.addWidget(card_widget)

        # Botão exportar
        export_rm_externos_btn = QPushButton("💾 Exportar RM Externos Fora do Prazo para Excel")
        export_rm_externos_btn.setObjectName("primaryButton")
        export_rm_externos_btn.clicked.connect(lambda: self.export_rm_externos_list(df_rm_externos))
        rm_externos_layout.addWidget(export_rm_externos_btn)

        # Tabela com lista de exames
        rm_externos_table = self.create_rm_externos_table(df_rm_externos)
        rm_externos_layout.addWidget(rm_externos_table)

        rm_externos_group.setLayout(rm_externos_layout)
        laudar_widgets.append(rm_externos_group)

        # Nova seção: Pacientes com RM Externos em Atraso (Agrupado)
        rm_pacientes_group = QGroupBox("👥 Pacientes com RMs Externas em Atraso - Visão Consolidada")
        rm_pacientes_layout = QVBoxLayout()

        # Agrupar por paciente
        if len(df_rm_externos) > 0:
            pacientes_agrupados = df_rm_externos.groupby(['SAME', 'NOME_PACIENTE']).agg({
                'TEMPO_ATENDIMENTO_MIN': ['count', 'mean', 'max'],
                'DATA_HORA_PRESCRICAO': 'min',
                'STATUS_ALAUDAR': 'max'
            }).reset_index()

            pacientes_agrupados.columns = ['SAME', 'NOME_PACIENTE', 'QTD_EXAMES', 'TEMPO_MEDIO', 'TEMPO_MAXIMO', 'PRIMEIRA_PRESCRICAO', 'ULTIMO_LAUDO']
            pacientes_agrupados = pacientes_agrupados.sort_values('QTD_EXAMES', ascending=False)

            # Cards com resumo por paciente
            cards_layout = QHBoxLayout()

            total_pacientes = len(pacientes_agrupados)
            pacientes_multiplos = len(pacientes_agrupados[pacientes_agrupados['QTD_EXAMES'] > 1])
            tempo_medio_geral = pacientes_agrupados['TEMPO_MEDIO'].mean()

            card_total_pac = StatCard(
                "Total de Pacientes",
                f"{total_pacientes}",
                "pacientes com RMs em atraso",
                "#FF6B6B"
            )

            card_multiplos = StatCard(
                "Pacientes com Múltiplos Exames",
                f"{pacientes_multiplos}",
                f"{(pacientes_multiplos/total_pacientes*100):.1f}% do total" if total_pacientes > 0 else "0%",
                "#FF9800"
            )

            card_tempo_medio = StatCard(
                "Tempo Médio por Paciente",
                f"{tempo_medio_geral:.0f} min",
                f"{(tempo_medio_geral/60):.1f} horas",
                "#F44336"
            )

            cards_layout.addWidget(card_total_pac)
            cards_layout.addWidget(card_multiplos)
            cards_layout.addWidget(card_tempo_medio)
            cards_layout.addStretch()

            cards_widget = QWidget()
            cards_widget.setLayout(cards_layout)
            rm_pacientes_layout.addWidget(cards_widget)

            # Botão exportar
            export_pacientes_btn = QPushButton("💾 Exportar Lista de Pacientes para Excel")
            export_pacientes_btn.setObjectName("primaryButton")
            export_pacientes_btn.clicked.connect(lambda: self.export_rm_pacientes_list(pacientes_agrupados))
            rm_pacientes_layout.addWidget(export_pacientes_btn)

            # Tabela de pacientes
            pacientes_table = self.create_rm_pacientes_table(pacientes_agrupados)
            rm_pacientes_layout.addWidget(pacientes_table)
        else:
            no_data_label = QLabel("✅ Nenhum paciente com RM externa em atraso")
            no_data_label.setStyleSheet("color: #51CF66; font-size: 14px; font-weight: bold; padding: 20px;")
            rm_pacientes_layout.addWidget(no_data_label)

        rm_pacientes_group.setLayout(rm_pacientes_layout)
        laudar_widgets.append(rm_pacientes_group)

        # Criar widgets para SLA Entrega
        entrega_widgets = []

        # Primeira linha: Distribuição SLA + Distribuição por Faixa de Tempo
        entrega_row1 = QWidget()
        entrega_row1_layout = QHBoxLayout(entrega_row1)
        chart_entrega_dist = self.create_entrega_distribution_chart(self.resultado['df_entrega'])
        chart_entrega_time_dist = self.create_entrega_time_distribution_chart(self.resultado['df_entrega'])
        entrega_row1_layout.addWidget(chart_entrega_dist)
        entrega_row1_layout.addWidget(chart_entrega_time_dist)
        entrega_widgets.append(entrega_row1)

        # Segunda linha: SLA por Modalidade + SLA por Tipo de Atendimento
        entrega_row2 = QWidget()
        entrega_row2_layout = QHBoxLayout(entrega_row2)
        chart_entrega_sla_mod = self.create_sla_entrega_by_modality_chart(self.resultado['analise_grupo_entrega'])
        chart_entrega_sla_type = self.create_entrega_sla_by_type_chart(self.resultado['analise_combinada_entrega'])
        entrega_row2_layout.addWidget(chart_entrega_sla_mod)
        entrega_row2_layout.addWidget(chart_entrega_sla_type)
        entrega_widgets.append(entrega_row2)

        # Terceira linha: Tempo Médio por Modalidade + Tempo Médio por Tipo
        entrega_row3 = QWidget()
        entrega_row3_layout = QHBoxLayout(entrega_row3)
        chart_entrega_avg_mod = self.create_avg_entrega_time_by_modality_chart(self.resultado['analise_grupo_entrega'])
        chart_entrega_avg_type = self.create_avg_entrega_time_by_type_chart(self.resultado['analise_combinada_entrega'])
        entrega_row3_layout.addWidget(chart_entrega_avg_mod)
        entrega_row3_layout.addWidget(chart_entrega_avg_type)
        entrega_widgets.append(entrega_row3)

        # Quarta linha: Heatmap combinado
        chart_entrega_heatmap = self.create_combined_entrega_chart(self.resultado['analise_combinada_entrega'])
        entrega_widgets.append(chart_entrega_heatmap)

        # ========== SEÇÃO DE ANÁLISE LONGITUDINAL MENSAL ==========
        # Título da seção
        longitudinal_title = QLabel("📈 Análise Longitudinal - Evolução Mensal por Modalidade")
        longitudinal_title.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #58a6ff;
                padding: 20px 10px 10px 10px;
                background-color: #161b22;
                border-radius: 8px;
                margin-top: 20px;
            }
        """)
        entrega_widgets.append(longitudinal_title)

        # Um gráfico por linha para melhor visualização
        # Linha 1: TC
        chart_long_tc = self.create_entrega_longitudinal_tc_chart(self.resultado['analise_longitudinal_entrega'])
        entrega_widgets.append(chart_long_tc)

        # Linha 2: RM
        chart_long_rm = self.create_entrega_longitudinal_rm_chart(self.resultado['analise_longitudinal_entrega'])
        entrega_widgets.append(chart_long_rm)

        # Linha 3: US
        chart_long_us = self.create_entrega_longitudinal_us_chart(self.resultado['analise_longitudinal_entrega'])
        entrega_widgets.append(chart_long_us)

        # Linha 4: RX
        chart_long_rx = self.create_entrega_longitudinal_rx_chart(self.resultado['analise_longitudinal_entrega'])
        entrega_widgets.append(chart_long_rx)

        # Linha 5: MN
        chart_long_mn = self.create_entrega_longitudinal_mn_chart(self.resultado['analise_longitudinal_entrega'])
        entrega_widgets.append(chart_long_mn)

        # ========== FIM DA SEÇÃO LONGITUDINAL ==========

        # ========== SEÇÃO DE SLA POR MODALIDADE E PORTA ==========
        # Título da seção
        sla_porta_title = QLabel("📊 SLA de Entrega por Modalidade e Porta (PA, Internado, Externo)")
        sla_porta_title.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #58a6ff;
                padding: 20px 10px 10px 10px;
                background-color: #161b22;
                border-radius: 8px;
                margin-top: 20px;
            }
        """)
        entrega_widgets.append(sla_porta_title)

        # Primeira linha: TC + RM
        sla_porta_row1 = QWidget()
        sla_porta_row1_layout = QHBoxLayout(sla_porta_row1)
        chart_sla_tc_porta = self.create_sla_by_modality_and_port_tc(self.resultado['analise_combinada_entrega'])
        chart_sla_rm_porta = self.create_sla_by_modality_and_port_rm(self.resultado['analise_combinada_entrega'])
        sla_porta_row1_layout.addWidget(chart_sla_tc_porta)
        sla_porta_row1_layout.addWidget(chart_sla_rm_porta)
        entrega_widgets.append(sla_porta_row1)

        # Segunda linha: US + RX
        sla_porta_row2 = QWidget()
        sla_porta_row2_layout = QHBoxLayout(sla_porta_row2)
        chart_sla_us_porta = self.create_sla_by_modality_and_port_us(self.resultado['analise_combinada_entrega'])
        chart_sla_rx_porta = self.create_sla_by_modality_and_port_rx(self.resultado['analise_combinada_entrega'])
        sla_porta_row2_layout.addWidget(chart_sla_us_porta)
        sla_porta_row2_layout.addWidget(chart_sla_rx_porta)
        entrega_widgets.append(sla_porta_row2)

        # Terceira linha: Mamografia + Medicina Nuclear
        sla_porta_row3 = QWidget()
        sla_porta_row3_layout = QHBoxLayout(sla_porta_row3)
        chart_sla_mama_porta = self.create_sla_by_modality_and_port_mama(self.resultado['analise_combinada_entrega'])
        chart_sla_mn_porta = self.create_sla_by_modality_and_port_mn(self.resultado['analise_combinada_entrega'])
        sla_porta_row3_layout.addWidget(chart_sla_mama_porta)
        sla_porta_row3_layout.addWidget(chart_sla_mn_porta)
        entrega_widgets.append(sla_porta_row3)

        # ========== FIM DA SEÇÃO DE SLA POR MODALIDADE E PORTA ==========

        # Tabela detalhada
        table_group_entrega = QGroupBox("Análise Detalhada: SLA de Entrega")
        table_layout_entrega = QVBoxLayout()
        table_entrega = self.create_entrega_table(self.resultado['analise_combinada_entrega'])
        table_layout_entrega.addWidget(table_entrega)
        table_group_entrega.setLayout(table_layout_entrega)
        entrega_widgets.append(table_group_entrega)

        # Lista de pacientes fora do SLA
        patient_list_entrega_group = QGroupBox("📋 Pacientes com Entrega Fora do Prazo")
        patient_list_entrega_layout = QVBoxLayout()

        export_entrega_btn = QPushButton("💾 Exportar Lista para Excel")
        export_entrega_btn.setObjectName("primaryButton")
        export_entrega_btn.clicked.connect(lambda: self.export_patient_entrega_list(self.resultado['df_entrega']))
        patient_list_entrega_layout.addWidget(export_entrega_btn)

        patient_entrega_table = self.create_patient_entrega_list_table(self.resultado['df_entrega'])
        patient_list_entrega_layout.addWidget(patient_entrega_table)
        patient_list_entrega_group.setLayout(patient_list_entrega_layout)
        entrega_widgets.append(patient_list_entrega_group)

        # Popolar dashboards
        self.dashboard_window.populate_laudar_dashboard(laudar_widgets)
        self.dashboard_window.populate_entrega_dashboard(entrega_widgets)
        convenio_widgets = self.create_convenio_section(self.resultado['df'])
        self.dashboard_window.populate_convenio_dashboard(convenio_widgets)

        # Mostrar janela
        self.dashboard_window.show()

    def generate_ai_analysis(self):
        """Gera análise estratégica com IA em janela separada"""
        if not self.resultado:
            QMessageBox.warning(self, "Aviso", "Execute a análise de dados primeiro!")
            return

        # Verificar configuração
        api_type = "OpenAI" if "OpenAI" in self.ai_type_combo.currentText() else "LM Studio"
        api_key = self.api_key_input.text().strip()
        api_url = self.api_url_input.text().strip()

        if api_type == "OpenAI" and not api_key:
            QMessageBox.warning(self, "Aviso", "Informe a API Key da OpenAI!")
            return

        if api_type == "LM Studio" and not api_url:
            api_url = "http://localhost:1234/v1"  # Default

        # Desabilitar botão
        self.generate_ai_btn.setEnabled(False)
        self.generate_ai_btn.setText("⏳ Gerando Análise...")

        # Criar e abrir janela de análise
        self.ai_window = AIAnalysisWindow(self)
        self.ai_window.set_loading()
        self.ai_window.show()

        # Iniciar thread
        self.ai_thread = AIAnalysisThread(
            self.resultado['stats_gerais'],
            self.resultado['stats_entrega'],
            self.resultado['analise_grupo'],
            self.resultado['analise_tipo'],
            api_type,
            api_key,
            api_url
        )
        self.ai_thread.finished.connect(self.on_ai_analysis_complete)
        self.ai_thread.error.connect(self.on_ai_analysis_error)
        self.ai_thread.start()

    def on_ai_analysis_complete(self, analise):
        """Callback quando análise IA completa"""
        if hasattr(self, 'ai_window'):
            self.ai_window.set_complete(analise)
        self.generate_ai_btn.setEnabled(True)
        self.generate_ai_btn.setText("🧠 Gerar Análise Estratégica")

    def on_ai_analysis_error(self, error_msg):
        """Callback quando erro na análise IA"""
        if hasattr(self, 'ai_window'):
            self.ai_window.set_error(error_msg)
        self.generate_ai_btn.setEnabled(True)
        self.generate_ai_btn.setText("🧠 Gerar Análise Estratégica")


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
