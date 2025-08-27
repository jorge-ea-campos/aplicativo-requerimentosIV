import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any
import logging
from dataclasses import dataclass
from enum import Enum
import traceback

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- Data Classes and Enums ---
class ProblemType(Enum):
    QR = "QR"  # Quebra de Requisito
    CH = "CH"  # Conflito de Horário
    OTHER = "OTHER"

class ParecerStatus(Enum):
    PENDENTE = "Pendente"
    DEFERIDO = "Deferido SG"
    INDEFERIDO = "Indeferido SG"
    COC_ANALYSIS = "Para análise COC."

@dataclass
class ColumnNames:
    """Centralized column names configuration"""
    NUSP = "nusp"
    PROBLEMA = "problema"
    PARECER = "parecer"
    NOME = "Nome completo"
    DISCIPLINA = "disciplina"
    ANO = "Ano"
    SEMESTRE = "Semestre"
    LINK = "link_requerimento"
    PLANO = "plano_estudo"
    PLANO_PRESENCA = "plano_presenca"
    PARECER_SG = "Parecer Serviço de Graduação"
    OBSERVACAO_SG = "Observação SG"

@dataclass
class FileRequirements:
    """File column requirements"""
    consolidado: List[str]
    requerimentos: List[str]

# --- Constants ---
COLS = ColumnNames()
REQUIRED_COLS = FileRequirements(
    consolidado=[COLS.NUSP, COLS.DISCIPLINA, COLS.ANO, COLS.SEMESTRE, COLS.PROBLEMA, COLS.PARECER],
    requerimentos=[COLS.NUSP, COLS.NOME, COLS.PROBLEMA, COLS.LINK, COLS.PLANO, COLS.PLANO_PRESENCA]
)

# --- Configuration ---
def configure_page():
    """Configure Streamlit page settings"""
    st.set_page_config(
        page_title="Sistema de Conferência de Requerimentos",
        page_icon="📋",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Initialize session state
    if 'decisions' not in st.session_state:
        st.session_state.decisions = {}
    if 'data_cache' not in st.session_state:
        st.session_state.data_cache = {}

def load_custom_css():
    """Load custom CSS styles"""
    st.markdown("""
    <style>
        .header-container {
            margin-bottom: 2rem;
        }
        .logo-and-title {
            display: flex;
            align-items: center;
            margin-bottom: 1rem;
        }
        .header-logo {
            height: 70px;
            margin-right: 20px;
        }
        .header-title-text {
            display: flex;
            flex-direction: column;
        }
        .university-name {
            font-size: 1.7rem;
            font-weight: bold;
            color: #333;
        }
        .department-name {
            font-size: 1.3rem;
            color: #0072b5;
        }
        .color-bar-yellow { height: 8px; background-color: #FDB913; }
        .color-bar-lightblue { height: 4px; background-color: #89cff0; }
        .color-bar-darkblue { height: 12px; background-color: #003366; }
        
        .main-header {
            font-size: 2.5rem; color: #1f77b4; text-align: center;
            padding: 1rem 0; border-bottom: 3px solid #1f77b4; margin-bottom: 2rem;
            margin-top: 0;
        }
        .metric-card {
            background-color: #f0f2f6; padding: 1.5rem; border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; margin-bottom: 1rem;
        }
        .error-box {
            background-color: #ffebee; 
            border-left: 4px solid #f44336; 
            padding: 1rem; 
            margin: 1rem 0;
        }
        .warning-box {
            background-color: #fff3e0; 
            border-left: 4px solid #ff9800; 
            padding: 1rem; 
            margin: 1rem 0;
        }
    </style>
    """, unsafe_allow_html=True)

# --- Data Loading and Processing ---
class DataLoader:
    """Handles data loading and validation"""
    
    @staticmethod
    def load_file(uploaded_file) -> Optional[pd.DataFrame]:
        """Load file with error handling"""
        try:
            # Try Excel first
            df = pd.read_excel(uploaded_file)
            logger.info(f"Successfully loaded {uploaded_file.name} as Excel")
            return df
        except Exception as excel_error:
            try:
                # Fall back to CSV
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file)
                logger.info(f"Successfully loaded {uploaded_file.name} as CSV")
                return df
            except Exception as csv_error:
                logger.error(f"Failed to load {uploaded_file.name}. Excel error: {excel_error}, CSV error: {csv_error}")
                st.error(f"Failed to read '{uploaded_file.name}'. Supported formats: Excel (.xlsx) or CSV (.csv)")
                return None

    @staticmethod
    def normalize_column_name(col_name: str) -> str:
        """Normalize column names for matching"""
        return str(col_name).lower().strip().replace(' ', '_')

    @classmethod
    def find_and_rename_columns(cls, df: pd.DataFrame, column_mapping: Dict[str, List[str]]) -> pd.DataFrame:
        """Find and rename columns based on mapping rules"""
        df_copy = df.copy()
        rename_dict = {}
        used_columns = set()
        
        # Normalize column mapping
        normalized_mapping = {}
        for target_col, possible_names in column_mapping.items():
            normalized_mapping[target_col] = [cls.normalize_column_name(name) for name in possible_names]
        
        # Find matches
        for col in df_copy.columns:
            if col in used_columns:
                continue
                
            normalized_col = cls.normalize_column_name(col)
            
            for target_col, normalized_names in normalized_mapping.items():
                if normalized_col in normalized_names and target_col not in rename_dict.values():
                    rename_dict[col] = target_col
                    used_columns.add(col)
                    break
        
        # Apply renaming
        df_copy.rename(columns=rename_dict, inplace=True)
        
        return df_copy

    @staticmethod
    def validate_required_columns(df: pd.DataFrame, required_cols: List[str], file_type: str) -> List[str]:
        """Validate that required columns exist"""
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            logger.warning(f"Missing columns in {file_type}: {missing_cols}")
        return missing_cols

    @staticmethod
    def clean_nusp_column(df: pd.DataFrame, file_name: str) -> pd.DataFrame:
        """Clean and validate NUSP column"""
        if COLS.NUSP not in df.columns:
            return df
            
        df_copy = df.copy()
        original_count = len(df_copy)
        
        # Convert to numeric
        df_copy[COLS.NUSP] = pd.to_numeric(df_copy[COLS.NUSP], errors='coerce')
        
        # Remove invalid NUSPs
        df_copy = df_copy.dropna(subset=[COLS.NUSP])
        df_copy[COLS.NUSP] = df_copy[COLS.NUSP].astype(int)
        
        invalid_count = original_count - len(df_copy)
        if invalid_count > 0:
            st.warning(f"Removed {invalid_count} records with invalid NUSP from {file_name}")
            
        return df_copy

# --- Analysis Functions ---
class DataAnalyzer:
    """Handles data analysis and metrics calculation"""
    
    @staticmethod
    def calculate_metrics(df_with_history: pd.DataFrame) -> Dict[str, Any]:
        """Calculate comprehensive metrics"""
        metrics = {}
        
        if df_with_history.empty:
            return metrics
        
        # Parecer analysis
        pareceres = df_with_history['parecer_historico'].str.lower().fillna('')
        aprovados = pareceres.str.contains('aprovado|deferido', na=False) & ~pareceres.str.contains('indeferido|negado', na=False)
        negados = pareceres.str.contains('indeferido|negado', na=False)
        
        total_com_parecer = aprovados.sum() + negados.sum()
        metrics['taxa_aprovacao'] = (aprovados.sum() / total_com_parecer * 100) if total_com_parecer > 0 else 0
        
        # Top disciplines
        if 'disciplina_historico' in df_with_history.columns:
            metrics['top_disciplinas'] = df_with_history['disciplina_historico'].value_counts().head(5)
        
        # Temporal distribution
        if all(col in df_with_history.columns for col in ['Ano_historico', 'Semestre_historico']):
            df_with_history['periodo'] = (
                df_with_history['Ano_historico'].astype(str) + '/' + 
                df_with_history['Semestre_historico'].astype(str)
            )
            metrics['distribuicao_temporal'] = df_with_history['periodo'].value_counts().sort_index()
        
        return metrics

# --- UI Components ---
class UIComponents:
    """Handles UI component rendering"""
    
    @staticmethod
    def render_header():
        """Render application header"""
        st.markdown("""
            <div class="header-container">
                <div class="logo-and-title">
                    <div class="header-title-text">
                        <span class="university-name">Universidade de São Paulo</span>
                        <span class="department-name">Serviço de Graduação - FZEA</span>
                    </div>
                </div>
                <div class="color-bar-container">
                    <div class="color-bar-yellow"></div>
                    <div class="color-bar-lightblue"></div>
                    <div class="color-bar-darkblue"></div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown('<h1 class="main-header">Sistema de Conferência de Requerimentos</h1>', unsafe_allow_html=True)

    @staticmethod
    def format_problem_type(problem: Any) -> str:
        """Format problem type with icons"""
        if pd.isna(problem):
            return "⚪ Não especificado"
        
        problem_str = str(problem).upper()
        if problem_str == "QR":
            return "🔴 Quebra de Requisito"
        elif problem_str == "CH":
            return "🟡 Conflito de Horário"
        else:
            return f"⚪ {problem}"

    @staticmethod
    def format_parecer(parecer: Any) -> str:
        """Format parecer with icons"""
        if pd.isna(parecer):
            return "📝 Pendente"
        
        parecer_str = str(parecer).lower()
        if any(word in parecer_str for word in ["negado", "indeferido"]):
            return f"❌ {parecer}"
        elif "aprovado" in parecer_str or "deferido" in parecer_str:
            return f"✅ {parecer}"
        else:
            return f"📝 {parecer}"

    @staticmethod
    def display_metrics(df_req: pd.DataFrame, df_with_history: pd.DataFrame, metrics: Dict[str, Any]):
        """Display main metrics"""
        st.markdown("### 📊 Métricas Principais")
        cols = st.columns(5)
        
        with cols[0]:
            st.metric("Total de Requerimentos", len(df_req))
        
        with cols[1]:
            alunos_unicos_hist = df_with_history[COLS.NUSP].nunique()
            total_alunos_req = df_req[COLS.NUSP].nunique()
            percentual_hist = (alunos_unicos_hist / total_alunos_req * 100) if total_alunos_req > 0 else 0
            st.metric("Alunos com Histórico", alunos_unicos_hist, f"{percentual_hist:.1f}%")
        
        with cols[2]:
            qr_count = (df_with_history.get("problema_historico", pd.Series()).str.upper() == "QR").sum()
            st.metric("Quebras de Requisito", qr_count)
        
        with cols[3]:
            ch_count = (df_with_history.get("problema_historico", pd.Series()).str.upper() == "CH").sum()
            st.metric("Conflitos de Horário", ch_count)
        
        with cols[4]:
            st.metric("Taxa de Aprovação", f"{metrics.get('taxa_aprovacao', 0):.1f}%")

    @staticmethod
    def display_charts(metrics: Dict[str, Any]):
        """Display analysis charts"""
        st.markdown("### 📈 Análise Gráfica")
        col1, col2 = st.columns(2)
        
        # Top disciplines chart
        if 'top_disciplinas' in metrics and not metrics['top_disciplinas'].empty:
            with col1:
                st.markdown("##### 📚 Top 5 Disciplinas")
                top_d = metrics['top_disciplinas']
                fig = px.bar(
                    x=top_d.values, 
                    y=top_d.index, 
                    orientation='h',
                    title="Disciplinas com Mais Pedidos"
                )
                fig.update_layout(
                    yaxis_title="Disciplina", 
                    xaxis_title="Número de Pedidos",
                    yaxis={'categoryorder': 'total ascending'}
                )
                st.plotly_chart(fig, use_container_width=True)
        
        # Temporal distribution chart
        if 'distribuicao_temporal' in metrics and not metrics['distribuicao_temporal'].empty:
            with col2:
                st.markdown("##### 🗓️ Pedidos por Período")
                dist_t = metrics['distribuicao_temporal']
                fig2 = px.line(
                    x=dist_t.index, 
                    y=dist_t.values, 
                    markers=True,
                    title="Distribuição Temporal de Pedidos"
                )
                fig2.update_layout(
                    xaxis_title="Período", 
                    yaxis_title="Número de Pedidos"
                )
                st.plotly_chart(fig2, use_container_width=True)

# --- Export Functions ---
class ExportHandler:
    """Handles data export functionality"""
    
    @staticmethod
    @st.cache_data
    def to_excel(df: pd.DataFrame) -> bytes:
        """Convert DataFrame to Excel format"""
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Relatorio')
            
            # Format worksheet
            worksheet = writer.sheets['Relatorio']
            header_format = writer.book.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BD',
                'border': 1
            })
            
            # Apply header formatting
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Auto-adjust column widths
            for i, col in enumerate(df.columns):
                max_length = max(
                    df[col].astype(str).map(len).max(),
                    len(col)
                ) + 2
                worksheet.set_column(i, i, min(max_length, 50))
        
        return output.getvalue()

    @staticmethod
    def prepare_export_data(df_req: pd.DataFrame, decisions: Dict[str, Dict[str, str]]) -> pd.DataFrame:
        """Apply decisions to DataFrame for export"""
        df_export = df_req.copy()
        
        # Ensure export columns exist
        for col in [COLS.PARECER_SG, COLS.OBSERVACAO_SG]:
            if col not in df_export.columns:
                df_export[col] = ""
        
        # Apply decisions
        for index, row in df_export.iterrows():
            decision_key = f"req_{index}"
            if decision_key in decisions:
                decision = decisions[decision_key]
                if decision['status'] != ParecerStatus.PENDENTE.value:
                    df_export.loc[index, COLS.PARECER_SG] = decision['status']
                    df_export.loc[index, COLS.OBSERVACAO_SG] = decision.get('justificativa', '')
        
        return df_export

# --- Main Application Logic ---
class RequestManagementApp:
    """Main application class"""
    
    def __init__(self):
        self.data_loader = DataLoader()
        self.analyzer = DataAnalyzer()
        self.ui = UIComponents()
        self.export_handler = ExportHandler()

    def setup_column_mapping(self) -> Tuple[Dict[str, List[str]], Dict[str, List[str]]]:
        """Define column mapping for both file types"""
        consolidado_mapping = {
            COLS.NUSP: ["nusp", "numero usp", "número usp", "n° usp", "n usp"],
            COLS.PROBLEMA: ["problema"],
            COLS.DISCIPLINA: ["disciplina"],
            COLS.ANO: ["ano"],
            COLS.SEMESTRE: ["semestre"],
            COLS.PARECER: ["parecer"]
        }
        
        requerimentos_mapping = {
            COLS.NUSP: ["nusp", "numero usp", "número usp", "n° usp", "n usp"],
            COLS.NOME: ["nome completo", "nome"],
            COLS.PROBLEMA: ["problema"],
            COLS.LINK: ["link para o requerimento", "links pedidos requerimento", "link_requerimento"],
            COLS.PLANO: ["plano de estudo", "link plano de estudos", "plano_estudo"],
            COLS.PLANO_PRESENCA: ["plano de presença", "link plano de presença", "plano_presenca"],
            COLS.OBSERVACAO_SG: ["observação sg"]
        }
        
        return consolidado_mapping, requerimentos_mapping

    def process_uploaded_files(self, file_consolidado, file_requerimentos, show_debug: bool = False) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
        """Process and validate uploaded files"""
        try:
            # Load files
            df_consolidado = self.data_loader.load_file(file_consolidado)
            df_requerimentos = self.data_loader.load_file(file_requerimentos)
            
            if df_consolidado is None or df_requerimentos is None:
                return None, None
            
            if show_debug:
                with st.expander("🔍 Debug - Original Columns"):
                    st.write("**Consolidado:**", df_consolidado.columns.tolist())
                    st.write("**Requerimentos:**", df_requerimentos.columns.tolist())
            
            # Get column mappings
            consolidado_mapping, requerimentos_mapping = self.setup_column_mapping()
            
            # Apply column mapping
            df_consolidado = self.data_loader.find_and_rename_columns(df_consolidado, consolidado_mapping)
            df_requerimentos = self.data_loader.find_and_rename_columns(df_requerimentos, requerimentos_mapping)
            
            # Validate required columns
            missing_consolidado = self.data_loader.validate_required_columns(
                df_consolidado, REQUIRED_COLS.consolidado, "consolidado"
            )
            missing_requerimentos = self.data_loader.validate_required_columns(
                df_requerimentos, REQUIRED_COLS.requerimentos, "requerimentos"
            )
            
            if missing_consolidado or missing_requerimentos:
                error_msg = []
                if missing_consolidado:
                    error_msg.append(f"Arquivo consolidado: colunas faltando - {', '.join(missing_consolidado)}")
                if missing_requerimentos:
                    error_msg.append(f"Arquivo requerimentos: colunas faltando - {', '.join(missing_requerimentos)}")
                raise ValueError("\n".join(error_msg))
            
            # Clean NUSP columns
            df_consolidado = self.data_loader.clean_nusp_column(df_consolidado, "consolidado")
            df_requerimentos = self.data_loader.clean_nusp_column(df_requerimentos, "requerimentos")
            
            # Rename historical columns
            hist_columns = {
                COLS.DISCIPLINA: f"{COLS.DISCIPLINA}_historico",
                COLS.ANO: f"{COLS.ANO}_historico",
                COLS.SEMESTRE: f"{COLS.SEMESTRE}_historico",
                COLS.PROBLEMA: f"{COLS.PROBLEMA}_historico",
                COLS.PARECER: f"{COLS.PARECER}_historico"
            }
            df_consolidado.rename(columns=hist_columns, inplace=True)
            df_requerimentos.rename(columns={COLS.PROBLEMA: 'problema_atual'}, inplace=True)
            
            return df_consolidado, df_requerimentos
            
        except Exception as e:
            st.error(f"Error processing files: {str(e)}")
            logger.error(f"File processing error: {traceback.format_exc()}")
            return None, None

    def display_student_details(self, df_requerimentos: pd.DataFrame, df_merged: pd.DataFrame):
        """Display interactive student details section"""
        st.markdown("### 📋 Análise de Requerimentos por Aluno")
        st.info("Clique no nome para ver o histórico e dar o parecer nos pedidos atuais.")
        
        # Get unique students
        alunos_unicos = (
            df_requerimentos[[COLS.NUSP, COLS.NOME]]
            .drop_duplicates(subset=[COLS.NUSP])
            .sort_values(COLS.NOME)
        )
        
        for _, aluno in alunos_unicos.iterrows():
            nusp_aluno = aluno[COLS.NUSP]
            
            with st.expander(f"👤 {aluno[COLS.NOME]} (NUSP: {nusp_aluno})"):
                self._render_student_requests(df_requerimentos, nusp_aluno)
                self._render_student_history(df_merged, nusp_aluno)

    def _render_student_requests(self, df_requerimentos: pd.DataFrame, nusp_aluno: int):
        """Render current requests for a student"""
        current_requests = df_requerimentos[df_requerimentos[COLS.NUSP] == nusp_aluno]
        
        st.markdown("##### 📌 Requerimento(s) no Semestre Atual")
        
        if current_requests.empty:
            st.write("Nenhum requerimento encontrado para este aluno.")
            return
        
        for index, request in current_requests.iterrows():
            decision_key = f"req_{index}"
            
            # Initialize decision if not exists
            if decision_key not in st.session_state.decisions:
                st.session_state.decisions[decision_key] = {
                    'status': ParecerStatus.PENDENTE.value,
                    'justificativa': ''
                }
            
            # Display request details
            problema_display = request.get('problema_atual', 'Não especificado')
            st.markdown(f"**Problema/Pedido:** `{problema_display}`")
            
            # Display links
            self._render_request_links(request)
            
            # Parecer selection
            self._render_parecer_section(decision_key)
            
            st.divider()

    def _render_request_links(self, request: pd.Series):
        """Render request links"""
        links = [
            (COLS.LINK, "🔗 Link para o Requerimento"),
            (COLS.PLANO, "📄 Link para o Plano de Estudo"),
            (COLS.PLANO_PRESENCA, "📋 Link para o Plano de Presença")
        ]
        
        for col, label in links:
            link = request.get(col, "")
            if pd.notna(link) and str(link).strip():
                st.markdown(f"**{label}:** [Acessar Link]({link})")
            else:
                st.markdown(f"**{label}:** Não informado")

    def _render_parecer_section(self, decision_key: str):
        """Render parecer selection section"""
        parecer_options = [status.value for status in ParecerStatus]
        current_status = st.session_state.decisions[decision_key]['status']
        
        if current_status not in parecer_options:
            current_status = ParecerStatus.PENDENTE.value
        
        # Status selection
        status = st.radio(
            "Parecer:",
            parecer_options,
            key=f"status_{decision_key}",
            index=parecer_options.index(current_status),
            horizontal=True
        )
        st.session_state.decisions[decision_key]['status'] = status
        
        # Justification input
        if status != ParecerStatus.PENDENTE.value:
            label_map = {
                ParecerStatus.DEFERIDO.value: "Justificativa para o deferimento:",
                ParecerStatus.INDEFERIDO.value: "Justificativa para o indeferimento:",
                ParecerStatus.COC_ANALYSIS.value: "Observações para o COC:"
            }
            label = label_map.get(status, "Justificativa (Opcional):")
            
            justificativa = st.text_area(
                label,
                value=st.session_state.decisions[decision_key]['justificativa'],
                key=f"just_input_{decision_key}"
            )
            
            if st.button("Salvar Justificativa", key=f"save_btn_{decision_key}"):
                st.session_state.decisions[decision_key]['justificativa'] = justificativa
                st.success("Justificativa salva!")
        else:
            st.session_state.decisions[decision_key]['justificativa'] = ''

    def _render_student_history(self, df_merged: pd.DataFrame, nusp_aluno: int):
        """Render student history"""
        st.markdown("##### 📜 Histórico de Pedidos de Requerimento")
        
        historico_aluno = df_merged[df_merged[COLS.NUSP] == nusp_aluno].copy()
        
        if historico_aluno.empty or historico_aluno['disciplina_historico'].isnull().all():
            st.info("Este aluno não possui histórico de pedidos anteriores.")
            return
        
        # Format history for display
        historico_aluno['problema_formatado'] = historico_aluno['problema_historico'].apply(
            self.ui.format_problem_type
        )
        historico_aluno['parecer_formatado'] = historico_aluno['parecer_historico'].apply(
            self.ui.format_parecer
        )
        
        # Select and rename columns for display
        display_cols = [
            'disciplina_historico', 'Ano_historico', 'Semestre_historico',
            'problema_formatado', 'parecer_formatado'
        ]
        
        df_display = historico_aluno[display_cols].rename(columns={
            col: col.replace('_historico', '').replace('_formatado', '').title() 
            for col in display_cols
        })
        
        st.dataframe(df_display, hide_index=True, use_container_width=True)

    def render_export_section(self, df_requerimentos: pd.DataFrame):
        """Render export section"""
        st.markdown("### 📥 Exportar Relatórios")
        
        # Prepare data with decisions
        df_com_pareceres = self.export_handler.prepare_export_data(
            df_requerimentos, st.session_state.decisions
        )
        
        # Filter non-rejected requests
        df_nao_indeferidos = df_com_pareceres[
            df_com_pareceres[COLS.PARECER_SG] != ParecerStatus.INDEFERIDO.value
        ].copy()
        
        # Export buttons
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("##### Relatório Completo com Pareceres")
            timestamp = datetime.now().strftime('%Y%m%d_%H%M')
            st.download_button(
                label="📥 Baixar como Excel",
                data=self.export_handler.to_excel(df_com_pareceres),
                file_name=f"relatorio_completo_pareceres_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    def render_sidebar(self) -> Tuple[Any, Any, bool]:
        """Render sidebar with file uploaders and settings"""
        with st.sidebar:
            st.header("📁 Upload de Arquivos")
            
            file_consolidado = st.file_uploader(
                "**1. Histórico de Pedidos (consolidado)**", 
                type=["xlsx", "csv"],
                help="Arquivo com histórico consolidado de pedidos anteriores"
            )
            
            file_requerimentos = st.file_uploader(
                "**2. Pedidos do Semestre Atual**", 
                type=["xlsx", "csv"],
                help="Arquivo com pedidos do semestre atual para análise"
            )
            
            st.info("💡 Os arquivos devem conter uma coluna com o número USP.")
            
            with st.expander("⚙️ Configurações Avançadas"):
                show_debug = st.checkbox("Mostrar informações de debug", value=False)
                
            return file_consolidado, file_requerimentos, show_debug

    def render_file_instructions(self):
        """Render file upload instructions"""
        st.markdown("### 🚀 Bem-vindo!")
        st.markdown("Para começar, faça o upload dos arquivos obrigatórios na barra lateral.")
        
        with st.expander("📋 Estrutura Esperada dos Arquivos"):
            st.markdown("**Arquivo Consolidado (Histórico):**")
            st.code(', '.join(REQUIRED_COLS.consolidado))
            
            st.markdown("**Arquivo de Requerimentos (Atual):**")
            st.code(', '.join(REQUIRED_COLS.requerimentos))
            
            st.markdown("""
            **Observações importantes:**
            - A coluna NUSP deve conter apenas números
            - Os links devem ser URLs válidas e acessíveis
            - Formatos suportados: Excel (.xlsx) e CSV (.csv)
            """)

    def run(self):
        """Main application runner"""
        try:
            # Setup
            configure_page()
            load_custom_css()
            
            # Render header
            self.ui.render_header()
            
            # Sidebar
            file_consolidado, file_requerimentos, show_debug = self.render_sidebar()
            
            # Check if files are uploaded
            if not (file_consolidado and file_requerimentos):
                self.render_file_instructions()
                return
            
            # Process files
            with st.spinner("Processando arquivos..."):
                df_consolidado, df_requerimentos = self.process_uploaded_files(
                    file_consolidado, file_requerimentos, show_debug
                )
                
                if df_consolidado is None or df_requerimentos is None:
                    st.error("Falha no processamento dos arquivos. Verifique os formatos e tente novamente.")
                    return
                
                # Merge data
                df_merged = df_requerimentos.merge(df_consolidado, on=COLS.NUSP, how="left")
                df_merged_with_history = df_merged.dropna(subset=['disciplina_historico'])
                
                # Calculate metrics
                metrics = self.analyzer.calculate_metrics(df_merged_with_history)
            
            # Display metrics
            self.ui.display_metrics(df_requerimentos, df_merged_with_history, metrics)
            st.divider()
            
            # Display charts if there's historical data
            if not df_merged_with_history.empty:
                self.ui.display_charts(metrics)
                st.divider()
            
            # Student details section
            self.display_student_details(df_requerimentos, df_merged)
            st.divider()
            
            # Export section
            self.render_export_section(df_requerimentos)
            
        except Exception as e:
            st.error(f"Erro inesperado na aplicação: {str(e)}")
            logger.error(f"Application error: {traceback.format_exc()}")
            
            if show_debug:
                st.exception(e)

# --- Authentication ---
class AuthenticationManager:
    """Handles user authentication"""
    
    @staticmethod
    def check_password() -> bool:
        """Check if user is authenticated"""
        if "password_correct" not in st.session_state:
            st.session_state["password_correct"] = False
        
        return st.session_state["password_correct"]
    
    @staticmethod
    def render_login_form() -> bool:
        """Render login form and handle authentication"""
        st.title("🔒 Acesso Restrito")
        
        try:
            correct_password = st.secrets["passwords"]["senha_mestra"]
        except (AttributeError, KeyError):
            st.error("❌ Aplicação não configurada. Contate o administrador.")
            st.info("""
            **Para desenvolvedores:** Configure a senha em `secrets.toml`:
            ```toml
            [passwords]
            senha_mestra = "sua_senha_aqui"
            ```
            """)
            return False
        
        with st.form("login_form"):
            password = st.text_input("Senha", type="password")
            submitted = st.form_submit_button("Entrar")
            
            if submitted:
                if password == correct_password:
                    st.session_state["password_correct"] = True
                    st.rerun()
                else:
                    st.error("❌ Senha incorreta.")
        
        return False

# --- Entry Point ---
def main():
    """Application entry point"""
    auth_manager = AuthenticationManager()
    
    # Check authentication
    if not auth_manager.check_password():
        auth_manager.render_login_form()
        return
    
    # Run main application
    app = RequestManagementApp()
    app.run()

if __name__ == "__main__":
    main()edocument.spreadsheetml.sheet"
            )
        
        with col2:
            st.markdown("##### Relatório de Pedidos Não Indeferidos")
            st.download_button(
                label="📥 Baixar como Excel",
                data=self.export_handler.to_excel(df_nao_indeferidos),
                file_name=f"relatorio_nao_indeferidos_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-offic

