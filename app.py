"""
人身险公司偿付能力信息分享平台  v5.0
  - 箱线图展示异常值/四分位/中位数
  - 分布直方图展示指标分布
  - 公司分类对比（大型/中型/小型/银行系/外资系/养老系）
  - 指标含义说明
  - 增资发债信息统计
"""
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import io
import os
from datetime import datetime
import json
import requests
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import base64

# PDF生成相关库
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor, white, black
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import warnings
import traceback
warnings.filterwarnings("ignore")

# ===========================================================
# 页面配置
# ===========================================================
st.set_page_config(
    page_title="人身险公司偿付能力信息分享平台",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ===========================================================
# CSS（PDF风格）
# ===========================================================
st.markdown("""
<style>
/* 登录页 */
.login-container {
    max-width: 480px; margin: 60px auto; padding: 40px 36px;
    background: transparent; border-radius: 16px;
    box-shadow: none;
}
.login-title { text-align:center; font-size:3.2rem !important; font-weight:800 !important; margin-bottom:8px !important; background: linear-gradient(90deg, #ffffff, #00d4ff); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; letter-spacing: 2px; }
.login-sub  { text-align:center; font-size:0.82rem !important; color:#a0c4ff !important; margin-bottom:32px !important; letter-spacing: 4px; text-transform: uppercase; }
.login-btn   { background:linear-gradient(135deg,#1a3a5c,#0d5fa5) !important; color:white !important;
                 border-radius:8px !important; font-size:1rem !important; padding:10px 0 !important; }
/* 主界面 */
section.main { background: #ffffff !important; }
.main-header {
    background: url(data:image/png;base64,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IECAwLJA8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoeWjskICBAgQIECAQI1A8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoZojt0YIECBAgAABAssCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhJaPygoJECBAgAABAjUCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhGqO3BohQIAAAQIECCwLFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEFo+KiskQIAAAQIECNQIFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEKo5cmuEAAECBAgQILAsULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ2j5qKyQAAECBAgQIFAjULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ6jmyK0RAgQIECBAgMCyQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKHlo7JCAgQIECBAgECNQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKGaI7dGCBAgQIAAAQLLAsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYSWj8oKCRAgQIAAAQI1AsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYRqjtwaIUCAAAECBAgsCxSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBBaPiorJECAAAECBAjUCBSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBCqOXJrhAABAgQIECCwLFC8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0No+aiskAABAgQIECBQI1C8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0Oo5sitEQIECBAgQIDAskDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyh5aOyQgIECBAgQIBAjUDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyhmiO3RggQIECAAAECywLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGElo/KCgkQIECAAAECNQLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGEao7cGiFAgAABAgQILAsUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQWj4qKyRAgAABAgQI1AgUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQqjlya4QAAQIECBAgsCxQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDaPmorJAAAQIECBAgUCNQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDqObIrRECBAgQIECAwLJA8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoeWjskICBAgQIECAQI1A8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoZojt0YIECBAgAABAssCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhJaPygoJECBAgAABAjUCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhGqO3BohQIAAAQIECCwLFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEFo+KiskQIAAAQIECNQIFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEKo5cmuEAAECBAgQILAsULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ2j5qKyQAAECBAgQIFAjULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ6jmyK0RAgQIECBAgMCyQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKHlo7JCAgQIECBAgECNQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKGaI7dGCBAgQIAAAQLLAsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYSWj8oKCRAgQIAAAQI1AsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYRqjtwaIUCAAAECBAgsCxSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBBaPiorJECAAAECBAjUCBSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBCqOXJrhAABAgQIECCwLFC8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0No+aiskAABAgQIECBQI1C8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0Oo5sitEQIECBAgQIDAskDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyh5aOyQgIECBAgQIBAjUDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyhmiO3RggQIECAAAECywLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGElo/KCgkQIECAAAECNQLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGEao7cGiFAgAABAgQILAsUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQWj4qKyRAgAABAgQI1AgUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQqjlya4QAAQIECBAgsCxQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDaPmorJAAAQIECBAgUCNQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDqObIrRECBAgQIECAwLJA8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoeWjskICBAgQIECAQI1A8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoZojt0YIECBAgAABAssCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhJaPygoJECBAgAABAjUCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhGqO3BohQIAAAQIECCwLFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLE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QIDAskDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyh5aOyQgIECBAgQIBAjUDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyhmiO3RggQIECAAAECywLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ9C4PWgAACAASURBVFVLgACBYgGxPDGElo/KCgkQIECAAAECNQLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGEao7cGiFAgAABAgQILAsUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQWj4qKyRAgAABAgQI1AgUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQqjlya4QAAQIECBAgsCxQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDaPmorJAAAQIECBAgUCNQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDqObIrRECBAgQIECAwLJA8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoeWjskICBAgQIECAQI1A8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoZojt0YIECBAgAABAssCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhJaPygoJECBAgAABAjUCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhGqO3BohQIAAAQIECCwLFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEFo+KiskQIAAAQIECNQIFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEKo5cmuEAAECBAgQILAsULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ2j5qKyQAAECBAgQIFAjULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ6jmyK0RAgQIECBAgMCyQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKHlo7JCAgQIECBAgECNQPFKrTUCBAhMBMRysZwA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AAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0No+aiskAABAgQIECBQI1C8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0Oo5sitEQIECBAgQIDAskDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyh5aOyQgIECBAgQIBAjUDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyhmiO3RggQIECAAAECywLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGElo/KCgkQIECAAAECNQLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGEao7cGiFAgAABAgQILAsUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQWj4qKyRAgAABAgQI1AgUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQqjlya4QAAQIECBAgsCxQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDaPmorJAAAQIECBAgUCNQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDqObIrRECBAgQIECAwLJA8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoeWjskICBAgQIECAQI1A8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoZojt0YIECBAgAABAssCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhJaPygoJECBAgAABAjUCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhGqO3BohQIAAAQIECCwLFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEFo+KiskQIAAAQIECNQIFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEKo5cmuEAAECBAgQILAsULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ2j5qKyQAAECBAgQIFAjULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ6jmyK0RAgQIECBAgMCyQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKHlo7JCAgQIECBAgECNQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKGaI7dGCBAgQIAAAQLLAsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYSWj8oKCRAgQIAAAQI1AsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYRqjtwaIUCAAAECBAgsCxSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBBaPiorJECAAAECBAjUCBSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBCqOXJrhAABAgQIECCwLFC8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0No+aiskAABAgQIECBQI1C8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0Oo5sitEQIECBAgQIDAskDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyh5aOyQgIECBAgQIBAjUDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyhmiO3RggQIECAAAECywLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGElo/KCgkQIECAAAECNQLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGEao7cGiFAgAABAgQILAsUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQWj4qKyRAgAABAgQI1AgUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQqjlya4QAAQIECBAgsCxQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDaPmorJAAAQIECBAgUCNQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDqObIrRECBAgQIECAwLJA8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoeWjskICBAgQIECAQI1A8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoZojt0YIECBAgAABAssCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhJaPygoJECBAgAABAjUCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhGqO3BohQIAAAQIECCwLFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEFo+KiskQIAAAQIECNQIFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEKo5cmuEAAECBAgQILAsULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ2j5qKyQAAECBAgQIFAjULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ6jmyK0RAgQIECBAgMCyQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKHlo7JCAgQIECBAgECNQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKGaI7dGCBAgQIAAAQLLAsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYSWj8oKCRAgQIAAAQI1AsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYRqjtwaIUCAAAECBAgsCxSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBBaPiorJECAAAECBAjUCBSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBCqOXJrhAABAgQIECCwLFC8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0No+aiskAABAgQIECBQI1C8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0Oo5sitEQIECBAgQIDAskDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyh5aOyQgIECBAgQIBAjUDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyhmiO3RggQIECAAAECywLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGElo/KCgkQIECAAAECNQLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGEao7cGiFAgAABAgQILAsUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQWj4qKyRAgAABAgQI1AgUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQqjlya4QAAQIECBAgsCxQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDaPmorJAAAQIECBAgUCNQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDqObIrRECBAgQIECAwLJA8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoeWjskICBAgQIECAQI1A8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoZojt0YIECBAgAABAssCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhJaPygoJECBAgAABAjUCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhGqO3BohQIAAAQIECCwLFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEFo+KiskQIAAAQIECNQIFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEKo5cmuEAAECBAgQILAsULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ2j5qKyQAAECBAgQIFAjULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ6jmyK0RAgQIECBAgMCyQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKHlo7JCAgQIECBAgECNQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKGaI7dGCBAgQIAAAQLLAsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYSWj8oKCRAgQIAAAQI1AsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYRqjtwaIUCAAAECBAgsCxSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBBaPiorJECAAAECBAjUCBSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBCqOXJrhAABAgQIECCwLFC8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0No+aiskAABAgQIECBQI1C8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0Oo5sitEQIECBAgQIDAskDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyh5aOyQgIECBAgQIBAjUDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyhmiO3RggQIECAAAECywLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGElo/KCgkQIECAAAECNQLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGEao7cGiFAgAABAgQILAsUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQWj4qKyRAgAABAgQI1AgUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQqjlya4QAAQIECBAgsCxQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDaPmorJAAAQIECBAgUCNQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDqObIrRECBAgQIECAwLJA8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoeWjskICBAgQIECAQI1A8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoZojt0YIECBAgAABAssCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhJaPygoJECBAgAABAjUCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhGqO3BohQIAAAQIECCwLFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEFo+KiskQIAAAQIECNQIFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEKo5cmuEAAECBAgQILAsULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ2j5qKyQAAECBAgQIFAjULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ6jmyK0RAgQIECBAgMCyQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKHlo7JCAgQIECBAgECNQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKGaI7dGCBAgQIAAAQLLAsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYSWKeGdNAAAC8lJREFUj8oKCRAgQIAAAQI1AsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYRqjtwaIUCAAAECBAgsCxSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBBaPiorJECAAAECBAjUCBSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBCqOXJrhAABAgQIECCwLFC8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0No+aiskAABAgQIECBQI1C8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0Oo5sitEQIECBAgQIDAskDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyh5aOyQgIECBAgQIBAjUDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyhmiO3RggQIECAAAECywLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGElo/KCgkQIECAAAECNQLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGEao7cGiFAgAABAgQILAsUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQWj4qKyRAgAABAgQI1AgUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQqjlya4QAAQIECBAgsCxQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDaPmorJAAAQIECBAgUCNQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDqObIrRECBAgQIECAwLJA8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoeWjskICBAgQIECAQI1A8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoZojt0YIECBAgAABAssCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhJaPygoJECBAgAABAjUCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhGqO3BohQIAAAQIECCwLFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEFo+KiskQIAAAQIECNQIFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEKo5cmuEAAECBAgQILAsULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ2j5qKyQAAECBAgQIFAjULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ6jmyK0RAgQIECBAgMCyQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKHlo7JCAgQIECBAgECNQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKGaI7dGCBAgQIAAAQLLAsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYSWj8oKCRAgQIAAAQI1AsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYRqjtwaIUCAAAECBAgsCxSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBBaPiorJECAAAECBAjUCBSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBCqOXJrhAABAgQIECCwLFC8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0No+aiskAABAgQIECBQI1C8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0Oo5sitEQIECBAgQIDAskDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyh5aOyQgIECBAgQIBAjUDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyhmiO3RggQIECAAAECywLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGElo/KCgkQIECAAAECNQLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGEao7cGiFAgAABAgQILAsUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQWj4qKyRAgAABAgQI1AgUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQqjlya4QAAQIECBAgsCxQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDaPmorJAAAQIECBAgUCNQvFJrjQABAhMBsVwsJ0CAAAECBAgQSAhMdla1BAgQKBYQyxNDqObIrRECBAgQIECAwLJA8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoeWjskICBAgQIECAQI1A8UqtNQIECEwExHKxnAABAgQIECBAICEw2VnVEiBAoFhALE8MoZojt0YIECBAgAABAssCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhJaPygoJECBAgAABAjUCxSu11ggQIDAREMvFcgIECBAgQIAAgYTAZGdVS4AAgWIBsTwxhGqO3BohQIAAAQIECCwLFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEFo+KiskQIAAAQIECNQIFK/UWiNAgMBEQCwXywkQIECAAAECBBICk51VLQECBIoFxPLEEKo5cmuEAAECBAgQILAsULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ2j5qKyQAAECBAgQIFAjULxSa40AAQITAbFcLCdAgAABAgQIEEgITHZWtQQIECgWEMsTQ6jmyK0RAgQIECBAgMCyQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKHlo7JCAgQIECBAgECNQPFKrTUCBAhMBMRysZwAAQIECBAgQCAhMNlZ1RIgQKBYQCxPDKGaI7dGCBAgQIAAAQLLAsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYSWj8oKCRAgQIAAAQI1AsUrtdYIECAwERDLxXICBAgQIECAAIGEwGRnVUuAAIFiAbE8MYRqjtwaIUCAAAECBAgsCxSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBBaPiorJECAAAECBAjUCBSv1FojQIDAREAsF8sJECBAgAABAgQSApOdVS0BAgSKBcTyxBCqOXJrhAABAgQIECCwLFC8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0No+aiskAABAgQIECBQI1C8UmuNAAECEwGxXCwnQIAAAQIECBBICEx2VrUECBAoFhDLE0Oo5sitEQIECBAgQIDAskDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyh5aOyQgIECBAgQIBAjUDxSq01AgQITATEcrGcAAECBAgQIEAgITDZWdUSIECgWEAsTwyhmiO3RggQIECAAAECywLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGElo/KCgkQIECAAAECNQLFK7XWCBAgMBEQy8VyAgQIECBAgACBhMBkZ1VLgACBYgGxPDGEao7cGiFAgAABAgQILAsUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigXE8sQQWj4qKyRAgAABAgQI1AgUr9RaI0CAwERALBfLCRAgQIAAAQIEEgKTnVUtAQIEigX+D9n4HqaYDvAeAAAAAElFTkSuQmCC);
    background-size: cover;
    background-position: right top;
    background-repeat: no-repeat;
    min-height: 120px;
    padding: 24px 28px;
    color: white;
    border-radius: 8px;
}
.main-header h1 { color: white; text-shadow: 0 1px 3px rgba(0,0,0,0.3); }
.main-header p { color: white; opacity: 0.9; }
.kpi-label { font-size: 0.78rem; color: #666; margin-bottom: 0.3rem; }
.kpi-value { font-size: 1.7rem; font-weight: 700; color: #1A3A5C; }
/* 指标含义卡片 */
.metric-explain {
    background: #f8f9fa; border-left: 4px solid #0d5fa5;
    padding: 6px 12px; border-radius: 0 8px 8px 0;
    margin: 6px 0; font-size: 0.88rem; color: #444; line-height: 1.6;
}
/* 权限标签 */
.role-badge-admin {
    background:#e8f4fd; color:#1a3a5c; padding:2px 10px;
    border-radius:12px; font-size:0.72rem; font-weight:600;
}
.role-badge-user {
    background:#f0f0f0; color:#666; padding:2px 10px;
    border-radius:12px; font-size:0.72rem; font-weight:600;
}
/* 导航高亮 */
div[data-test-id="stRadio"] label[data-selected="true"] {
    background:#e8f4fd !important; border-left:3px solid #0d5fa5 !important;
}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
[data-test-id="stSidebar"] { min-width: 300px !important; }
section[data-test-id="stSidebar"] .stMultiSelect [data-baseweb="tag"] { font-size: 0.7rem; }

/* ========== 打印样式 ========== */
@media print {
    /* 隐藏侧边栏、顶部导航、按钮等 —— 使用多重选择器确保生效 */
    [data-testid="stSidebar"],
    [data-test-id="stSidebar"],
    section[data-testid="stSidebar"],
    .css-1d391kg,
    .stSidebar {
        display: none !important;
        width: 0 !important;
        min-width: 0 !important;
        max-width: 0 !important;
        overflow: hidden !important;
        position: absolute !important;
        left: -9999px !important;
    }
    header, footer, #MainMenu { display: none !important; }
    .print-buttons-area, .print-hint { display: none !important; }
    [data-test-id="stToolbar"], .stDeployButton { display: none !important; }

    /* 主内容区占满宽度（侧边栏隐藏后） */
    .main .block-container {
        max-width: 100% !important;
        padding: 10mm !important;
        margin: 0 !important;
        width: 100% !important;
    }
    .main {
        padding-left: 0 !important;
        margin-left: 0 !important;
    }

    /* 强制分页：section-break 在打印时生效 */
    .section-break {
        page-break-after: always !important;
        break-after: page !important;
        height: 0 !important;
        margin: 0 !important;
        border: none !important;
    }

    /* 避免图表、表格、指标卡片被截断 */
    .js-plotly-plot,
    .stDataFrame,
    .stTable,
    table,
    [data-testid="stMetric"],
    [data-testid="stMetricValue"],
    [data-testid="stMetricLabel"],
    .element-container {
        page-break-inside: avoid !important;
        break-inside: avoid !important;
    }

    /* 强制 expander 内容在打印时可见 */
    details,
    details[open],
    [data-testid="stExpander"] details,
    [data-testid="stExpander"] details summary + div {
        display: block !important;
        visibility: visible !important;
        height: auto !important;
        overflow: visible !important;
    }
    /* 隐藏 expander 的折叠/展开箭头（打印时不必要） */
    details summary {
        font-weight: bold;
        margin-bottom: 8px;
    }

    /* st.dataframe 滚动容器在打印时展开 */
    .stDataFrame > div:first-child {
        overflow: visible !important;
        height: auto !important;
        max-height: none !important;
    }
    .stDataFrame table {
        overflow: visible !important;
    }
    /* 强制并列栏在打印时不换行 */
    [data-test-id="stColumns"],
    [data-test-id="stColumns"] {
        display: flex !important;
        flex-wrap: nowrap !important;
    }
    [data-test-id="stColumns"] > [data-test-id="stColumn"],
    [data-test-id="stColumns"] > [data-test-id="stColumn"] {
        flex: 1 0 50% !important;
        min-width: 50% !important;
    }
    /* 打印时缩小箱型图X轴标签字体，防止截断 */
    .xtick text, g[class*="xtick"] text {
        font-size: 8px !important;
    }
}

/* 屏幕显示时 section-break 不可见 */
.section-break {
    height: 1px;
    margin: 0;
    border: none;
}
/* 统一页面主标题字号 */
h3 { font-size: 1.4rem !important; font-weight: 700 !important; margin-top: 0.1rem !important; margin-bottom: 0.1rem !important; }
/* 统一大标题字号 */
h4 { font-size: 1.1rem !important; font-weight: 600 !important; margin-top: 0rem !important; margin-bottom: 0rem !important; }
/* 统一小标题字号 */
h5 { font-size: 1.0rem !important; font-weight: 550 !important; margin-top: 0rem !important; margin-bottom: 0rem !important; }
/* 调整 st.info() 目标公司背景框高度 */
[data-testid="stAlert"] {
    padding: 16px 20px !important;
    margin: -4px 0px 0px 0px !important;
    min-height: 48px !important;
}
[data-testid="stAlert"] > div:first-child {
    padding: 8px 4px !important;
}
[data-testid="stAlert"] p {
    margin: 0 !important;
    line-height: 1.8 !important;
}

</style>
""", unsafe_allow_html=True)

# ===========================================================
# 常量 & 配置
# ===========================================================
SHEET_NAME = "2025Q4"
DEFAULT_Q   = "2025Q4"

RATIO_COMP = "综合偿付能力充足率"
RATIO_CORE = "核心偿付能力充足率"
COMP_PCT    = "综合偿付能力充足率(%)"
CORE_PCT    = "核心偿付能力充足率(%)"
THR_COMP_PASS = 1.0   # 100%
THR_COMP_WARN = 1.5   # 150%

# 公司分类（与PDF一致）
# 指标分类（65个指标）
IND_CATS = {
    "01-核心偿付能力指标": [RATIO_COMP, RATIO_CORE, "综合偿付能力溢额","核心偿付能力溢额"],
    "02-资本规模指标": ["实际资本","最低资本","认可资产","认可负债","核心一级资本","核心二级资本","附属一级资本","附属二级资本"],
    "03-资本效率比率": ["核心资本/注册资本","实际资本/注册资本","认可资产/注册资本","总资产/注册资本",
                          "实际资本/认可资产","净利润/实际资本"],
    "04-保单未来盈余": ["计入核心一级资本的保单未来盈余","计入核心二级资本的保单未来盈余",
                          "保单未来盈余/核心资本","保单未来盈余/保险合同负债"],
    "05-量化风险最低资本": ["量化风险最低资本","寿险业务保险风险最低资本合计",
                             "市场风险-最低资本合计","信用风险-最低资本合计"],
    "06-保险风险": ["寿险业务保险风险最低资本合计","非寿险业务保险风险最低资本合计"],
    "07-市场风险": ["市场风险-最低资本合计","市场风险-利率风险最低资本","市场风险-权益价格风险最低资本",
                     "市场风险-房地产价格风险最低资本","市场风险-汇率风险最低资本"],
    "08-信用风险": ["信用风险-最低资本合计","信用风险-利差风险最低资本","信用风险-交易对手违约风险最低资本"],
}

LOGIN_PASSWORD = "KPMG1234"
ADMIN_USERS    = {"admin"}
NORMAL_USERS   = {"user", "guest"}

# ===========================================================
# 工具函数
# ===========================================================
def load_comp_cats_from_df(df):
    """从DataFrame中动态读取公司分类（A列=分类，B列=公司）"""
    comp_cats = {}
    if "分类" in df.columns and "公司" in df.columns:
        for _, row in df.iterrows():
            cat = str(row["分类"]).strip()
            co = str(row["公司"]).strip()
            if cat and co and cat != "nan" and co != "nan":
                if cat not in comp_cats:
                    comp_cats[cat] = []
                if co not in comp_cats[cat]:
                    comp_cats[cat].append(co)
    return comp_cats

def load_data(file_path, sheet_name):
    """读取Excel，第0行为列名（分类/公司/各指标）"""
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
    return df

def clean_df(df):
    df = df.copy()
    # 确保分类和公司列存在
    if "分类" not in df.columns or "公司" not in df.columns:
        # 如果列名不对，尝试修复
        if df.columns[0] != "分类":
            df.rename(columns={df.columns[0]: "分类"}, inplace=True)
        if df.columns[1] != "公司":
            df.rename(columns={df.columns[1]: "公司"}, inplace=True)

    df["公司"] = df["公司"].astype(str).str.strip()
    df = df[df["公司"].str.len() > 0]
    df = df[~df["公司"].isin(["nan","NaN","None","合计","小计"])]

    # 将所有数值列转为数字
    for col in df.columns:
        if col not in ("分类","公司"):
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df.reset_index(drop=True)

def pct(v):
    try:
        return f"{float(v)*100:.1f}%"
    except:
        return "—"

def fmt_wan(v):
    try:
        v = float(v)
        if abs(v) >= 10000:
            return f"{v/10000:.1f}亿"
        return f"{v:,.0f}万"
    except:
        return "—"

def hex_to_rgba(hex_color, alpha):
    """将 #RRGGBB 转为 rgba(R,G,B,A) 字符串"""
    hex_color = hex_color.lstrip("#")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


# ===========================================================
# 页面渲染函数（用于支持一键显示全部）
# ===========================================================

def render_page_01(standalone=True):
    """01-行业整体偿付能力概览"""
    # 加大标题段前间隔（打印时生效）
    st.markdown('<div style="margin-top: 16px;"></div>', unsafe_allow_html=True)
    st.markdown(f"### 🏦 01 · 行业整体偿付能力概览 · {q_label}")

    # 总览内容下移一点
    st.markdown('<div style="margin-top: 40px;"></div>', unsafe_allow_html=True)

    if has_c:
        med  = df[RATIO_COMP].median()
        n_p  = (df[RATIO_COMP] >= THR_COMP_WARN).sum()
        n_w  = ((df[RATIO_COMP]>=THR_COMP_PASS)&(df[RATIO_COMP]<THR_COMP_WARN)).sum()
        n_f  = (df[RATIO_COMP] < THR_COMP_PASS).sum()

        # 第一行：主要指标（大号）
        kcols_main = st.columns(3)
        kcols_main[0].metric("纳入公司数", f"{len(df)} 家")
        kcols_main[1].metric("综合偿付能力充足率中位数", pct(med))
        if has_k:
            med_k = df[RATIO_CORE].median()
            kcols_main[2].metric("核心偿付能力充足率中位数", pct(med_k))
        else:
            kcols_main[2].metric("充足（≥150%）", f"{n_p} 家",  delta=f"{n_p/len(df)*100:.0f}%")

        # 第二行：分类统计
        kcols_sub = st.columns(3)
        kcols_sub[0].metric("充足（≥150%）", f"{n_p} 家",  delta=f"{n_p/len(df)*100:.0f}%")
        kcols_sub[1].metric("预警（100-150%）", f"{n_w} 家", delta=f"{n_w}家" if n_w>0 else "无", delta_color="off")
        kcols_sub[2].metric("不达标（<100%）", f"{n_f} 家", delta=f"⚠{n_f}家" if n_f>0 else "无", delta_color="inverse")

    elif has_k:
        med_k = df[RATIO_CORE].median()
        kcols = st.columns(2)
        kcols[0].metric("纳入公司数", f"{len(df)} 家")
        kcols[1].metric("核心偿付能力充足率中位数", pct(med_k))

    st.divider()
    st.markdown('<div style="page-break-after: always;"></div>', unsafe_allow_html=True)

    # 公司分类列表（参照PDF第5页）
    st.markdown("##### 🏢 调研公司分类列表")

    # 分类标准说明（紧凑样式）
    st.markdown("""
    <div style="font-size:0.95rem; color:#666; margin: 2px 0 6px 0; line-height:1.4;">
        <strong>分类标准：</strong>认可资产规模大于5000亿的属于大型公司，小于500亿的属于小型公司，介于中间的是中型公司。
    </div>
    """, unsafe_allow_html=True)

    # 构建纵向表格（参考KPMG PDF格式）
    cat_names = ["大型公司", "中型公司", "小型公司", "银行系", "外资系", "养老健康"]
    cat_data = {}
    max_len = 0
    for cat in cat_names:
        cos = st.session_state.get("comp_cats", {}).get(cat, [])
        df_cat = df[df["公司"].isin(cos)]
        names = df_cat["公司"].tolist() if not df_cat.empty else []
        cat_data[cat] = names
        max_len = max(max_len, len(names))

    # 生成HTML表格（紧凑样式，确保单页可放下）
    html_rows = []
    for row_idx in range(max_len):
        cells = [f'<td style="border:1px solid #ddd; padding:1px 4px; text-align:center; font-size:0.82rem; color:#555; width:36px;">{row_idx + 1}</td>']
        for cat in cat_names:
            names = cat_data[cat]
            name = names[row_idx] if row_idx < len(names) else ""
            cells.append(f'<td style="border:1px solid #ddd; padding:1px 6px; text-align:center; font-size:0.82rem; color:#333; min-width:80px;">{name}</td>')
        html_rows.append("<tr>" + "".join(cells) + "</tr>")

    header_cells = ['<th style="border:1px solid #ddd; padding:2px 4px; text-align:center; font-size:0.85rem; font-weight:600; color:#1A3A5C; background:#f0f4f8; width:60px;">序号</th>']
    for cat in cat_names:
        header_cells.append(f'<th style="border:1px solid #ddd; padding:2px 6px; text-align:center; font-size:0.85rem; font-weight:600; color:#1A3A5C; background:#f0f4f8; min-width:80px;">{cat}</th>')

    table_html = f"""
    <div style="overflow-x:auto; margin-bottom:12px;">
    <table style="border-collapse:collapse; width:100%; border:1px solid #ddd;">
        <thead>
            <tr>{"".join(header_cells)}</tr>
        </thead>
        <tbody>
            {"".join(html_rows)}
        </tbody>
    </table>
    </div>
    """
    st.markdown(table_html, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)  # 关闭不可分区块

    st.divider()
    st.markdown('<div style="page-break-after: always;"></div>', unsafe_allow_html=True)

    # 偿付能力充足率整体分布情况（统一标题 + 两列并排）
    st.markdown("#### 📊 偿付能力充足率整体分布情况")
    st.markdown("""
    <div style="font-size:1.05rem; color:#444; margin: 4px 0 12px 0; line-height:1.5;">
        我国71家人身险公司的综合及核心偿付能力充足率呈现较为集中的分布，将人身险公司按照特征类型及认可资产规模进行分类，六大类公司的充足率分布如下图。
    </div>
    """, unsafe_allow_html=True)

    box_cols = st.columns(2)

    # 综合偿付能力充足率（左列）
    with box_cols[0]:
        st.markdown("##### 📈 综合偿付能力充足率")
        if target_co and COMP_PCT in df.columns:
            t_mask = df["公司"] == target_co
            if t_mask.any():
                t_val = df.loc[t_mask, COMP_PCT].iloc[0]
                st.info(f"🎯 **目标公司：{target_co}** — 综合偿付能力充足率：**{t_val:.1f}%**")

        if has_c:
            fig_box = boxplot_with_annotations(
                df, COMP_PCT, yaxis_title="综合偿付能力充足率(%)",
                height=360, target_co=target_co, y_multiplier=1.0,
                group_by_category=True, y_max=450
            )
            if fig_box:
                fig_box.add_hline(y=100, line_dash="dash", line_color="#E24B4A", annotation_text="100% 监管下限", annotation_font_color="#E24B4A")
                st.plotly_chart(fig_box, use_container_width=True)

    # 核心偿付能力充足率（右列）
    with box_cols[1]:
        st.markdown("##### 📈 核心偿付能力充足率")
        if target_co and CORE_PCT in df.columns:
            t_mask = df["公司"] == target_co
            if t_mask.any():
                t_val = df.loc[t_mask, CORE_PCT].iloc[0]
                st.info(f"🎯 **目标公司：{target_co}** — 核心偿付能力充足率：**{t_val:.1f}%**")

        if has_k:
            fig_box_k = boxplot_with_annotations(
                df, CORE_PCT, yaxis_title="核心偿付能力充足率(%)",
                height=360, target_co=target_co, y_multiplier=1.0,
                group_by_category=True, y_max=450
            )
            if fig_box_k:
                fig_box_k.add_hline(y=50, line_dash="dash", line_color="#E24B4A", annotation_text="50% 监管下限", annotation_font_color="#E24B4A")
                st.plotly_chart(fig_box_k, use_container_width=True)

    st.divider()
    st.markdown('<div style="page-break-after: always;"></div>', unsafe_allow_html=True)

    # 资本使用效率相关指标（统一标题 + 两列并排）
    st.markdown("#### 📊 资本使用效率相关指标")
    st.markdown("""
    <div style="font-size:1.05rem; color:#444; margin: 4px 0 16px 0; line-height:1.6;">
        核心资本/注册资本率说明股东投入资本的资本使用效率，部分小公司的该指标已低于100%，考虑是否因历史业务存在明显亏损导致；<br>
        实际资本/认可资产率较高的公司，说明目前的综合偿付能力充足率较高，但也说明还有提升资本利用率的空间。
    </div>
    """, unsafe_allow_html=True)

    cap_cols = st.columns(2)

    # 核心资本/注册资本率分布（左列）
    with cap_cols[0]:
        st.markdown("##### 📈 核心资本/注册资本率")
        if target_co and "核心资本/注册资本" in df.columns:
            t_mask = df["公司"] == target_co
            if t_mask.any():
                t_val = df.loc[t_mask, "核心资本/注册资本"].iloc[0]
                st.info(f"🎯 **目标公司：{target_co}** — 核心资本/注册资本率：**{t_val*100:.1f}%**")

        if "核心资本/注册资本" in df.columns:
            fig_box_core_reg = boxplot_with_annotations(
                df, "核心资本/注册资本",
                yaxis_title="核心资本/注册资本率",
                height=360, target_co=target_co, y_multiplier=1.0,
                group_by_category=True, y_min=-1, y_max=30, pct_display=True, dtick=5
            )
            if fig_box_core_reg:
                fig_box_core_reg.add_hline(y=1.0, line_dash="dash", line_color="#E24B4A", annotation_text="100%", annotation_font_color="#E24B4A")
                st.plotly_chart(fig_box_core_reg, use_container_width=True)

    # 实际资本/认可资产率分布（右列）
    with cap_cols[1]:
        st.markdown("##### 📈 实际资本/认可资产率")
        if target_co and "实际资本/认可资产" in df.columns:
            t_mask = df["公司"] == target_co
            if t_mask.any():
                t_val = df.loc[t_mask, "实际资本/认可资产"].iloc[0]
                st.info(f"🎯 **目标公司：{target_co}** — 实际资本/认可资产率：**{t_val*100:.1f}%**")

        if "实际资本/认可资产" in df.columns:
            fig_box_ac_aa = boxplot_with_annotations(
                df, "实际资本/认可资产",
                yaxis_title="实际资本/认可资产率",
                height=360, target_co=target_co, y_multiplier=1.0,
                group_by_category=True, y_max=0.6, pct_display=True
            )
            if fig_box_ac_aa:
                st.plotly_chart(fig_box_ac_aa, use_container_width=True)

    st.divider()

# —— 02-实际资本数据对比分析 ——

def render_page_02(standalone=True):
    """02-实际资本数据对比分析"""
    if not standalone:
        st.markdown('<div style="page-break-after: always;"></div>', unsafe_allow_html=True)
    st.markdown(f"### 💰 02 · 实际资本数据对比分析 · {q_label}")

    # 资本分级行业整体情况
    st.markdown("#### 📊 资本分级行业整体情况")

    # 计算行业平均比例（用于说明文字）
    if all(c in df.columns for c in ["核心一级资本", "核心二级资本", "附属一级资本", "附属二级资本", "实际资本"]):
        core_pct = ((df["核心一级资本"] + df["核心二级资本"]) / df["实际资本"]).median()
        aff_pct  = ((df["附属一级资本"] + df["附属二级资本"]) / df["实际资本"]).median()
        st.markdown(f"""
        <div class="metric-explain">
            • <strong>核心资本</strong>是指在持续经营状态下和破产清算状态下均可以吸收损失的资本，<strong>附属资本</strong>是指在破产清算状态下可以吸收损失的资本；<br>
            • 行业平均核心资本占实际资本的比例约<strong>{core_pct*100:.0f}%</strong>，附属资本占比约<strong>{aff_pct*100:.0f}%</strong>。
        </div>
        """, unsafe_allow_html=True)

        # 资本分级行业分布箱线图
        st.markdown("##### 📈 资本分级行业分布情况")
        fig_cap = capital_tier_boxplot(df, target_co=target_co, height=360)
        fig_cap.update_layout(margin=dict(r=110))
        fig_cap.update_xaxes(range=[-0.5, 3.8])
        if fig_cap:
            st.plotly_chart(fig_cap, use_container_width=True)

    st.divider()
    st.markdown('<div class="section-break"></div>', unsafe_allow_html=True)

    # 核心资本占比分布情况
    st.markdown("#### 📊 核心资本占比分布情况")

    # 计算占比列（用于箱线图）
    df["核心一级资本占比"] = df["核心一级资本"] / df["实际资本"]
    df["核心二级资本占比"] = df["核心二级资本"] / df["实际资本"]

    st.markdown("""
    <div class="metric-explain">
        • 大型公司和银行系的核心一级资本占比较为相似，55%至70%；<br>
        • 核心二级资本占比较高的公司，通常有优先股、财务再或保单盈余的影响。
    </div>
    """, unsafe_allow_html=True)

    # 两个箱线图并排
    core_cols = st.columns(2)

    with core_cols[0]:
        st.markdown("##### 📈 核心一级资本占比")
        # 显示目标公司数值
        if target_co and "核心一级资本占比" in df.columns:
            t_mask = df["公司"] == target_co
            if t_mask.any():
                t_val = df.loc[t_mask, "核心一级资本占比"].iloc[0]
                st.info(f"🎯 **目标公司：{target_co}** — 核心一级资本占比：**{t_val*100:.0f}%**")

        fig_core1 = boxplot_with_annotations(
            df, "核心一级资本占比",
            yaxis_title="核心一级资本占比",
            height=360, target_co=target_co, y_multiplier=1.0,
            group_by_category=True, y_min=0.3, y_max=1.0, pct_display=True
        )
        if fig_core1:
            st.plotly_chart(fig_core1, use_container_width=True)

    with core_cols[1]:
        st.markdown("##### 📈 核心二级资本占比")
        # 显示目标公司数值
        if target_co and "核心二级资本占比" in df.columns:
            t_mask = df["公司"] == target_co
            if t_mask.any():
                t_val = df.loc[t_mask, "核心二级资本占比"].iloc[0]
                st.info(f"🎯 **目标公司：{target_co}** — 核心二级资本占比：**{t_val*100:.0f}%**")

        fig_core2 = boxplot_with_annotations(
            df, "核心二级资本占比",
            yaxis_title="核心二级资本占比",
            height=360, target_co=target_co, y_multiplier=1.0,
            group_by_category=True, y_max=0.3, pct_display=True
        )
        if fig_core2:
            st.plotly_chart(fig_core2, use_container_width=True)

    st.divider()
    st.markdown('<div class="section-break"></div>', unsafe_allow_html=True)

    # 附属资本占比分布情况
    st.markdown("#### 📊 附属资本占比分布情况")

    # 计算占比列（用于箱线图）
    df["附属一级资本占比"] = df["附属一级资本"] / df["实际资本"]
    df["附属二级资本占比"] = df["附属二级资本"] / df["实际资本"]

    st.markdown("""
    <div class="metric-explain">
        • 大型公司的附属一级和二级资本占实际资本的比例相对其他类型公司，分布较为集中，中小型和养老健康类公司的占比较为分散；<br>
        • 附属二级资本占比较高的公司，考虑是否有财务再合同的影响及保单未来盈余的分布情况。
    </div>
    """, unsafe_allow_html=True)

    # 两个箱线图并排
    aff_cols = st.columns(2)

    with aff_cols[0]:
        st.markdown("##### 📈 附属一级资本占比")
        # 显示目标公司数值
        if target_co and "附属一级资本占比" in df.columns:
            t_mask = df["公司"] == target_co
            if t_mask.any():
                t_val = df.loc[t_mask, "附属一级资本占比"].iloc[0]
                st.info(f"🎯 **目标公司：{target_co}** — 附属一级资本占比：**{t_val*100:.0f}%**")

        fig_aff1 = boxplot_with_annotations(
            df, "附属一级资本占比",
            yaxis_title="附属一级资本占比",
            height=360, target_co=target_co, y_multiplier=1.0,
            group_by_category=True, y_max=0.5, pct_display=True
        )
        if fig_aff1:
            st.plotly_chart(fig_aff1, use_container_width=True)

    with aff_cols[1]:
        st.markdown("##### 📈 附属二级资本占比")
        # 显示目标公司数值
        if target_co and "附属二级资本占比" in df.columns:
            t_mask = df["公司"] == target_co
            if t_mask.any():
                t_val = df.loc[t_mask, "附属二级资本占比"].iloc[0]
                st.info(f"🎯 **目标公司：{target_co}** — 附属二级资本占比：**{t_val*100:.0f}%**")

        fig_aff2 = boxplot_with_annotations(
            df, "附属二级资本占比",
            yaxis_title="附属二级资本占比",
            height=360, target_co=target_co, y_multiplier=1.0,
            group_by_category=True, y_max=0.15, pct_display=True
        )
        if fig_aff2:
            st.plotly_chart(fig_aff2, use_container_width=True)

    st.divider()

    st.markdown('<div class="section-break"></div>', unsafe_allow_html=True)

    # 资本分级情况：保单未来盈余
    st.markdown("#### 📊 资本分级情况：保单未来盈余")

    # 计算计入核心资本的保单未来盈余/核心资本比例
    df["计入核心资本的保单未来盈余"] = df["计入核心一级资本的保单未来盈余"] + df["计入核心二级资本的保单未来盈余"]
    df["核心资本"] = df["核心一级资本"] + df["核心二级资本"]
    df["计入核心资本的保单未来盈余占比"] = df["计入核心资本的保单未来盈余"] / df["核心资本"]

    # 动态统计
    over_40_count = int((df["计入核心资本的保单未来盈余占比"] > 0.40).sum())
    le_zero_count = int((df["计入核心资本的保单未来盈余占比"] <= 0).sum())

    st.markdown(f"""
    <div class="metric-explain">
        • 保单未来盈余是指保险公司现行有效寿险保单剩余期限所对应的当期确认的实际资本。<br>
        • 人身险行业本季度末有<strong>{over_40_count}家</strong>公司计入核心资本的保单未来盈余占核心资本的比例超过监管规定的40%的比例限制。
        另外，有<strong>{le_zero_count}家</strong>公司的保单未来盈余小于等于0%。
    </div>
    """, unsafe_allow_html=True)

    # 两个图表并排
    pfl_cols = st.columns(2)

    with pfl_cols[0]:
        st.markdown("##### 📈 计入核心资本的保单未来盈余/核心资本的比例")
        # 按区间统计（合并≤0%）
        import numpy as np
        bins = [-np.inf, 0.0001, 0.20, 0.35, 0.40, 0.55, np.inf]
        labels = ["≤0%", "(0,20%]", "(20%,35%]", "(35%,40%]", "(40%,55%]", "大于55%"]
        df["保单未来盈余区间"] = pd.cut(df["计入核心资本的保单未来盈余占比"], bins=bins, labels=labels, include_lowest=True)
        counts = df["保单未来盈余区间"].value_counts().sort_index()

        # 颜色映射
        color_map = {
            "≤0%": "#E24B4A",
            "(0,20%]": "#2EB872",
            "(20%,35%]": "#2E7AD6",
            "(35%,40%]": "#7B5DBE",
            "(40%,55%]": "#E8913A",
            "大于55%": "#34495E",
        }

        fig_bar = go.Figure()
        for label in labels:
            if label in counts.index:
                fig_bar.add_trace(go.Bar(
                    x=[label], y=[counts[label]],
                    marker_color=color_map.get(label, "#2E7AD6"),
                    text=[str(counts[label])],
                    textposition="outside",
                    textfont=dict(size=13, family="SimHei"),
                    showlegend=False,
                ))

        fig_bar.update_layout(
            height=360,
            margin=dict(l=40, r=20, t=20, b=60),
            plot_bgcolor="white", paper_bgcolor="white",
            yaxis_title="公司数量",
            yaxis=dict(gridcolor="#e0e0e0", zeroline=False, title_font=dict(size=13, family="SimHei")),
            xaxis=dict(tickfont=dict(size=10, family="SimHei"), tickangle=-15),  # X轴刻度10px
            bargap=0.3,
        )
        st.plotly_chart(fig_bar, use_container_width=True)

        # 备注
        st.caption("注：多数公司未披露计入核心二级资本和附属一级和二级资本的保单未来盈余，未披露的按照对应计入保单未来盈余为0来统计。"
                   "因此计入核心资本的保单未来盈余占核心资本的比例应较此处统计数据更高。")

    with pfl_cols[1]:
        st.markdown("##### 📈 计入核心资本的保单未来盈余/核心资本（箱线图）")
        # 显示目标公司数值
        if target_co and "计入核心资本的保单未来盈余占比" in df.columns:
            t_mask = df["公司"] == target_co
            if t_mask.any():
                t_val = df.loc[t_mask, "计入核心资本的保单未来盈余占比"].iloc[0]
                st.info(f"🎯 **目标公司：{target_co}** — 计入核心资本的保单未来盈余占比：**{t_val*100:.0f}%**")

        fig_pfl_box = boxplot_with_annotations(
            df, "计入核心资本的保单未来盈余占比",
            yaxis_title="计入核心资本的保单未来盈余占比",
            height=360, target_co=target_co, y_multiplier=1.0,
            group_by_category=True, y_min=-0.3, y_max=0.7, pct_display=True
        )
        if fig_pfl_box:
            # 添加40%监管红线
            fig_pfl_box.add_hline(
                y=0.40, line_dash="dash", line_color="red", line_width=1.5,
                annotation_text="监管限制 40%", annotation_position="right",
                annotation_font=dict(size=11, color="red", family="SimHei")
            )
            st.plotly_chart(fig_pfl_box, use_container_width=True)

    st.divider()

    st.markdown('<div class="section-break"></div>', unsafe_allow_html=True)

    # 核心一级/附属一级资本中保单未来盈余占比情况
    st.markdown("#### 📊 核心一级/附属一级资本中保单未来盈余占比情况")

    st.markdown("""
    <div class="metric-explain">
        • 多家公司未披露计入附属一级资本的保单未来盈余，因此附属一级资本中的保单未来盈余比例很多都是0%。
    </div>
    """, unsafe_allow_html=True)

    # 两个箱线图并排
    tier_pfl_cols = st.columns(2)

    with tier_pfl_cols[0]:
        st.markdown("##### 📈 核心一级资本中的保单未来盈余比例")
        # 显示目标公司数值
        if target_co and "核心一级资本中的保单未来盈余占比" in df.columns:
            t_mask = df["公司"] == target_co
            if t_mask.any():
                t_val = df.loc[t_mask, "核心一级资本中的保单未来盈余占比"].iloc[0]
                st.info(f"🎯 **目标公司：{target_co}** — 核心一级资本中的保单未来盈余比例：**{t_val*100:.0f}%**")

        fig_tier1 = boxplot_with_annotations(
            df, "核心一级资本中的保单未来盈余占比",
            yaxis_title="核心一级资本中的保单未来盈余比例",
            height=360, target_co=target_co, y_multiplier=1.0,
            group_by_category=True, y_min=-0.4, y_max=0.8, pct_display=True
        )
        if fig_tier1:
            st.plotly_chart(fig_tier1, use_container_width=True)

    with tier_pfl_cols[1]:
        st.markdown("##### 📈 附属一级资本中的保单未来盈余比例")
        # 显示目标公司数值
        if target_co and "附属一级资本中保单未来盈余占比" in df.columns:
            t_mask = df["公司"] == target_co
            if t_mask.any():
                t_val = df.loc[t_mask, "附属一级资本中保单未来盈余占比"].iloc[0]
                st.info(f"🎯 **目标公司：{target_co}** — 附属一级资本中的保单未来盈余比例：**{t_val*100:.0f}%**")

        fig_tier2 = boxplot_with_annotations(
            df, "附属一级资本中保单未来盈余占比",
            yaxis_title="附属一级资本中的保单未来盈余比例",
            height=360, target_co=target_co, y_multiplier=1.0,
            group_by_category=True, y_min=-0.2, y_max=1.0, pct_display=True
        )
        if fig_tier2:
            st.plotly_chart(fig_tier2, use_container_width=True)

    st.divider()
    st.markdown('<div class="section-break"></div>', unsafe_allow_html=True)

    # 保单未来盈余/保险合同负债
    if "保单未来盈余/保险合同负债" in df.columns:
        # 计算行业平均（总体比例 = 总保单未来盈余 / 总保险合同负债）
        total_pfl = (df["计入核心一级资本的保单未来盈余"] + df["计入核心二级资本的保单未来盈余"] +
                     df["计入附属一级资本的保单未来盈余"] + df["计入附属二级资本的保单未来盈余"]).sum()
        total_liability = df["保险合同负债"].sum()
        avg_pfl = total_pfl / total_liability if total_liability != 0 else 0

        st.markdown("#### 📊 保单未来盈余/保险合同负债（存量保单盈利能力）")
        st.markdown(f"""
        <div class="metric-explain">
            • 保单未来盈余/保险合同负债代表存量保单盈利能力，该指标越高代表公司长期负债盈利能力越强；<br>
            • 头部大型公司的存量保单盈利能力表现高于大部分的其他中小类型险企；<br>
            • 本季度末行业平均水平{avg_pfl*100:.1f}%。
        </div>
        """, unsafe_allow_html=True)

        # 显示目标公司数值
        if target_co and "保单未来盈余/保险合同负债" in df.columns:
            t_mask = df["公司"] == target_co
            if t_mask.any():
                t_val = df.loc[t_mask, "保单未来盈余/保险合同负债"].iloc[0]
                st.info(f"🎯 **目标公司：{target_co}** — 保单未来盈余/保险合同负债：**{t_val*100:.1f}%**")

        # 箱线图
        fig_pfl = boxplot_with_annotations(
            df, "保单未来盈余/保险合同负债",
            yaxis_title="保单未来盈余/保险合同负债",
            height=360, target_co=target_co, y_multiplier=1.0,
            group_by_category=True, y_min=-0.1, y_max=0.3, pct_display=True
        )
        fig_pfl.update_layout(margin=dict(r=150))
        fig_pfl.update_xaxes(range=[-0.5, 6.5])
        if fig_pfl:
            st.plotly_chart(fig_pfl, use_container_width=True)

        # 备注
        st.caption("注：保险合同负债为会计口径负债加独立账户负债。")



def render_page_03(standalone=True):
    """03-最低资本数据对比分析"""
    if not standalone:
        st.markdown('<div style="page-break-after: always;"></div>', unsafe_allow_html=True)
    st.markdown(f"### ⚠️ 03 · 最低资本数据对比分析 · {q_label}")

    st.markdown("""
    <div class="metric-explain">
        • 人身险行业本季度末最低资本以市场风险为主，保险风险平均占比高于信用风险；<br>
        • 行业的风险分散效应和损失吸收都较为集中，有个别公司表现突出。
    </div>
    """, unsafe_allow_html=True)

    # 量化风险最低资本 构成饼图 + 箱型图
    qc_col = "量化风险最低资本"
    if qc_col in df.columns:
        st.markdown("#### 📊 量化风险最低资本构成（行业合计）")
        # 两列布局：左饼图，右箱型图
        col_pie, col_box = st.columns([3, 7])

        with col_pie:
            risk_cols = [c for c in ["寿险业务保险风险最低资本合计","非寿险业务保险风险最低资本合计","市场风险-最低资本合计","信用风险-最低资本合计"] if c in df.columns]
            if risk_cols:
                sums = {c.replace("最低资本合计","").replace("-",""): abs(df[c].sum()) for c in risk_cols}
                fig_pie = px.pie(
                    names=list(sums.keys()),
                    values=list(sums.values()),
                    color_discrete_sequence=["#185FA5","#2E7AD6","#EF9F27","#1D9E75"]
                )
                fig_pie.update_layout(height=360, margin=dict(l=0,r=0,t=10,b=0))
                st.plotly_chart(fig_pie, use_container_width=True)

        with col_box:
            # 最低资本构成分布箱型图
            mc_indicators = [
                {"col": "寿险业务保险风险最低资本合计", "label": "保险风险（寿）", "color": "#185FA5"},
                {"col": "非寿险业务保险风险最低资本合计", "label": "保险风险（非寿）", "color": "#2E7AD6"},
                {"col": "市场风险-最低资本合计", "label": "市场风险", "color": "#EF9F27"},
                {"col": "信用风险-最低资本合计", "label": "信用风险", "color": "#1D9E75"},
                {"col": "量化风险分散效应", "label": "风险分散效应", "color": "#7B5DBE"},
                {"col": "特定类别保险合同损失吸收效应", "label": "损失吸收", "color": "#E8913A"},
                {"col": "考虑特征系数影响", "label": "特征系数影响", "color": "#E24B4A"},
                {"col": "控制风险最低资本", "label": "控制风险", "color": "#34495E"},
            ]
            # 过滤掉不存在的列
            mc_indicators = [ind for ind in mc_indicators if ind["col"] in df.columns and "最低资本" in df.columns]

            if mc_indicators:
                # 箱型图固定为全行业结果，不随公司类型筛选变化
                fig_mc = mc_composition_boxplot(
                    df_all, mc_indicators, "最低资本",
                    target_co=target_co, height=360
                )
                fig_mc.update_layout(margin=dict(r=110))
                fig_mc.update_xaxes(range=[-0.5, 7.8])
                if fig_mc:
                    st.plotly_chart(fig_mc, use_container_width=True)

        st.markdown('<div class="section-break"></div>', unsafe_allow_html=True)
        # 保险风险最低资本情况
        st.markdown("#### 📊 保险风险最低资本情况")
        st.markdown("""
        <div class="metric-explain">
            • 寿险保险风险占比较高的主要是小型公司、外资系公司和养老健康类公司，且分布较为分散；<br>
            • 非寿险保险风险占比明显较高的主要是养老健康类公司。
        </div>
        """, unsafe_allow_html=True)

        # 两列布局：左保险风险（寿），右保险风险（非寿）
        col_life, col_nonlife = st.columns(2)

        with col_life:
            st.markdown("##### 📈 寿险业务保险风险最低资本占比")
            # 显示目标公司数值（移到箱型图上方）
            if target_co and "寿险业务保险风险最低资本合计" in df.columns and "最低资本" in df.columns:
                t_mask = df["公司"] == target_co
                if t_mask.any():
                    t_row = df[t_mask].iloc[0]
                    if pd.notna(t_row["最低资本"]) and t_row["最低资本"] != 0:
                        t_val = t_row["寿险业务保险风险最低资本合计"] / t_row["最低资本"]
                        if pd.notna(t_val):
                            st.info(f"🎯 **目标公司：{target_co}** — 寿险业务保险风险最低资本占比：**{t_val*100:.1f}%**")
            fig_life = boxplot_with_annotations(
                df, "寿险业务保险风险最低资本合计/最低资本",
                "寿险业务保险风险最低资本占比",
                height=360, target_co=target_co,
                group_by_category=True, y_min=0, y_max=0.8, pct_display=True
            )
            if fig_life:
                st.plotly_chart(fig_life, use_container_width=True)

        with col_nonlife:
            st.markdown("##### 📈 非寿险业务保险风险最低资本占比")
            # 显示目标公司数值（移到箱型图上方）
            if target_co and "非寿险业务保险风险最低资本合计" in df.columns and "最低资本" in df.columns:
                t_mask = df["公司"] == target_co
                if t_mask.any():
                    t_row = df[t_mask].iloc[0]
                    if pd.notna(t_row["最低资本"]) and t_row["最低资本"] != 0:
                        t_val = t_row["非寿险业务保险风险最低资本合计"] / t_row["最低资本"]
                        if pd.notna(t_val):
                            st.info(f"🎯 **目标公司：{target_co}** — 非寿险业务保险风险最低资本占比：**{t_val*100:.1f}%**")
            fig_nonlife = boxplot_with_annotations(
                df, "非寿险业务保险风险最低资本合计/最低资本",
                "非寿险业务保险风险最低资本占比",
                height=360, target_co=target_co,
                group_by_category=True, y_min=0, y_max=0.08, pct_display=True
            )
            if fig_nonlife:
                st.plotly_chart(fig_nonlife, use_container_width=True)
        st.markdown('<div class="section-break"></div>', unsafe_allow_html=True)
        # 市场和信用风险最低资本情况
        st.markdown("#### 📊 市场和信用风险最低资本情况")
        st.markdown("""
        <div class="metric-explain">
            • 市场风险占比较高的公司类型分布较为均匀，大型公司和银行系的市场风险占比较低；<br>
            • 信用风险占比较高的公司主要是小型公司和外资系公司。
        </div>
        """, unsafe_allow_html=True)

        # 两列布局：左市场风险，右信用风险
        col_market, col_credit = st.columns(2)

        with col_market:
            st.markdown("##### 📈 市场风险最低资本占比")
            # 显示目标公司数值（移到箱型图上方）
            if target_co and "市场风险-最低资本合计" in df.columns and "最低资本" in df.columns:
                t_mask = df["公司"] == target_co
                if t_mask.any():
                    t_row = df[t_mask].iloc[0]
                    if pd.notna(t_row["最低资本"]) and t_row["最低资本"] != 0:
                        t_val = t_row["市场风险-最低资本合计"] / t_row["最低资本"]
                        if pd.notna(t_val):
                            st.info(f"🎯 **目标公司：{target_co}** — 市场风险最低资本占比：**{t_val*100:.1f}%**")
            fig_market = boxplot_with_annotations(
                df, "市场风险-最低资本合计/最低资本",
                "市场风险最低资本占比",
                height=360, target_co=target_co,
                group_by_category=True, y_min=0.2, y_max=1.2, pct_display=True
            )
            if fig_market:
                st.plotly_chart(fig_market, use_container_width=True)

        with col_credit:
            st.markdown("##### 📈 信用风险最低资本占比")
            # 显示目标公司数值（移到箱型图上方）
            if target_co and "信用风险-最低资本合计" in df.columns and "最低资本" in df.columns:
                t_mask = df["公司"] == target_co
                if t_mask.any():
                    t_row = df[t_mask].iloc[0]
                    if pd.notna(t_row["最低资本"]) and t_row["最低资本"] != 0:
                        t_val = t_row["信用风险-最低资本合计"] / t_row["最低资本"]
                        if pd.notna(t_val):
                            st.info(f"🎯 **目标公司：{target_co}** — 信用风险最低资本占比：**{t_val*100:.1f}%**")
            fig_credit = boxplot_with_annotations(
                df, "信用风险-最低资本合计/最低资本",
                "信用风险最低资本占比",
                height=360, target_co=target_co,
                group_by_category=True, y_min=0, y_max=1.0, pct_display=True
            )
            if fig_credit:
                st.plotly_chart(fig_credit, use_container_width=True)

        st.markdown('<div class="section-break"></div>', unsafe_allow_html=True)
        # 市场风险最低资本占认可资产率
        st.markdown("#### 📊 市场风险最低资本占认可资产率")
        # 市场风险最低资本占认可资产率
        st.markdown("""
        <div class="metric-explain">
            • 市场风险由利率风险、权益价格风险、房地产价格风险、境外固收/权益价格风险、汇率风险、风险分散效应构成；<br>
            • 利率风险、权益价格风险是市场风险的主要组成部分。
        </div>
        """, unsafe_allow_html=True)

        # 两列布局：左利率风险/认可资产率，右权益价格风险/认可资产率
        col_rate, col_equity = st.columns(2)

        with col_rate:
            st.markdown("##### 📈 利率风险/认可资产率")
            # 显示目标公司数值（箱型图上方）
            if target_co and "市场风险-利率风险最低资本" in df.columns and "认可资产" in df.columns:
                t_mask = df["公司"] == target_co
                if t_mask.any():
                    t_row = df[t_mask].iloc[0]
                    if pd.notna(t_row["认可资产"]) and t_row["认可资产"] != 0:
                        t_val = t_row["市场风险-利率风险最低资本"] / t_row["认可资产"]
                        if pd.notna(t_val):
                            st.info(f"🎯 **目标公司：{target_co}** — 利率风险/认可资产率：**{t_val*100:.1f}%**")
            fig_rate = boxplot_with_annotations(
                df, "利率风险/认可资产",
                "利率风险/认可资产率",
                height=360, target_co=target_co,
                group_by_category=True, y_min=0, y_max=0.1, pct_display=True
            )
            if fig_rate:
                st.plotly_chart(fig_rate, use_container_width=True)

        with col_equity:
            st.markdown("##### 📈 权益价格风险/认可资产率")
            # 显示目标公司数值（箱型图上方）
            if target_co and "市场风险-权益价格风险最低资本" in df.columns and "认可资产" in df.columns:
                t_mask = df["公司"] == target_co
                if t_mask.any():
                    t_row = df[t_mask].iloc[0]
                    if pd.notna(t_row["认可资产"]) and t_row["认可资产"] != 0:
                        t_val = t_row["市场风险-权益价格风险最低资本"] / t_row["认可资产"]
                        if pd.notna(t_val):
                            st.info(f"🎯 **目标公司：{target_co}** — 权益价格风险/认可资产率：**{t_val*100:.1f}%**")
            fig_equity = boxplot_with_annotations(
                df, "权益价格风险/认可资产",
                "权益价格风险/认可资产率",
                height=360, target_co=target_co,
                group_by_category=True, y_min=0, y_max=0.1, pct_display=True
            )
            if fig_equity:
                st.plotly_chart(fig_equity, use_container_width=True)

        st.markdown('<div class="section-break"></div>', unsafe_allow_html=True)
        # 信用风险最低资本占认可资产率
        st.markdown("#### 📊 信用风险最低资本占认可资产率")
        st.markdown("""
        <div class="metric-explain">
            • 信用风险由利差风险、交易对手违约风险等构成；
            • 利差风险是信用风险的主要组成部分。
        </div>
        """, unsafe_allow_html=True)

        # 两列布局：左利差风险/认可资产率，右对手违约风险/认可资产率
        col_spread, col_default = st.columns(2)

        with col_spread:
            st.markdown("##### 📈 利差风险/认可资产率")
            # 显示目标公司数值（箱型图上方）
            if target_co and "信用风险-利差风险最低资本" in df.columns and "认可资产" in df.columns:
                t_mask = df["公司"] == target_co
                if t_mask.any():
                    t_row = df[t_mask].iloc[0]
                    if pd.notna(t_row["认可资产"]) and t_row["认可资产"] != 0:
                        t_val = t_row["信用风险-利差风险最低资本"] / t_row["认可资产"]
                        if pd.notna(t_val):
                            st.info(f"🎯 **目标公司：{target_co}** — 利差风险/认可资产率：**{t_val*100:.1f}%**")
            fig_spread = boxplot_with_annotations(
                df, "利差风险/认可资产",
                "利差风险/认可资产率",
                height=360, target_co=target_co,
                group_by_category=True, y_min=0, y_max=0.04, pct_display=True
            )
            if fig_spread:
                st.plotly_chart(fig_spread, use_container_width=True)

        with col_default:
            st.markdown("##### 📈 对手违约风险/认可资产率")
            # 显示目标公司数值（箱型图上方）
            if target_co and "信用风险-交易对手违约风险最低资本" in df.columns and "认可资产" in df.columns:
                t_mask = df["公司"] == target_co
                if t_mask.any():
                    t_row = df[t_mask].iloc[0]
                    if pd.notna(t_row["认可资产"]) and t_row["认可资产"] != 0:
                        t_val = t_row["信用风险-交易对手违约风险最低资本"] / t_row["认可资产"]
                        if pd.notna(t_val):
                            st.info(f"🎯 **目标公司：{target_co}** — 对手违约风险/认可资产率：**{t_val*100:.1f}%**")
            fig_default = boxplot_with_annotations(
                df, "对手违约风险/认可资产",
                "对手违约风险/认可资产率",
                height=360, target_co=target_co,
                group_by_category=True, y_min=0, y_max=0.04, pct_display=True
            )
            if fig_default:
                st.plotly_chart(fig_default, use_container_width=True)
        st.markdown('<div class="section-break"></div>', unsafe_allow_html=True)

        # 风险分散效应和损失吸收
        st.markdown("#### 📊 风险分散效应和损失吸收")

        # 动态生成说明文字：根据各分类的中位数计算
        disp_desc = ""
        absrp_desc = ""
        if "量化风险分散效应/最低资本" in df.columns:
            disp_medians = {}
            for cat, cos in st.session_state.get("comp_cats", {}).items():
                vals = df[df["公司"].isin(cos)]["量化风险分散效应/最低资本"].dropna()
                if len(vals) > 0:
                    disp_medians[cat] = float(vals.median())
            # 风险分散效应为负值，中位数越小（越负）说明效应越强
            sorted_disp = sorted(disp_medians.items(), key=lambda x: x[1])
            if len(sorted_disp) >= 2:
                weakest = sorted_disp[-1]  # 中位数最大（最接近0）的分类
                weak_cats = [cat for cat, v in sorted_disp if v >= weakest[1] - 0.02]
                disp_desc = f"{'、'.join(weak_cats)}公司的风险分散效应稍弱"
            elif len(sorted_disp) == 1:
                disp_desc = f"{sorted_disp[0][0]}公司的风险分散效应稍弱"

        if "特定类别保险合同损失吸收效应/最低资本" in df.columns:
            absrp_medians = {}
            for cat, cos in st.session_state.get("comp_cats", {}).items():
                vals = df[df["公司"].isin(cos)]["特定类别保险合同损失吸收效应/最低资本"].dropna()
                if len(vals) > 0:
                    absrp_medians[cat] = float(vals.median())
            # 损失吸收效应为负值，中位数越小（越负）说明效应越明显
            sorted_absrp = sorted(absrp_medians.items(), key=lambda x: x[1])
            if len(sorted_absrp) >= 2:
                strongest = sorted_absrp[0]  # 中位数最小（最负）的分类
                strong_cats = [cat for cat, v in sorted_absrp if v <= strongest[1] + 0.02]
                absrp_desc = f"{'、'.join(strong_cats)}公司的损失吸收效应较为明显"
            elif len(sorted_absrp) == 1:
                absrp_desc = f"{sorted_absrp[0][0]}公司的损失吸收效应较为明显"

        desc_html = "<div class=\"metric-explain\">\n"
        if disp_desc:
            desc_html += f"    • {disp_desc}；<br>\n"
        if absrp_desc:
            desc_html += f"    • {absrp_desc}。\n"
        desc_html += "</div>"
        st.markdown(desc_html, unsafe_allow_html=True)

        # 两列布局：左风险分散效应，右损失吸收
        col_disp, col_absrp = st.columns(2)

        with col_disp:
            st.markdown("##### 📈 风险分散效应最低资本占比")
            # 显示目标公司数值（箱型图上方）
            if target_co and "量化风险分散效应" in df.columns and "最低资本" in df.columns:
                t_mask = df["公司"] == target_co
                if t_mask.any():
                    t_row = df[t_mask].iloc[0]
                    if pd.notna(t_row["最低资本"]) and t_row["最低资本"] != 0:
                        t_val = t_row["量化风险分散效应"] / t_row["最低资本"]
                        if pd.notna(t_val):
                            st.info(f"🎯 **目标公司：{target_co}** — 风险分散效应最低资本占比：**{t_val*100:.1f}%**")
            fig_disp = boxplot_with_annotations(
                df, "量化风险分散效应/最低资本",
                "风险分散效应最低资本占比",
                height=360, target_co=target_co,
                group_by_category=True, y_min=-0.6, y_max=0.0, pct_display=True
            )
            if fig_disp:
                st.plotly_chart(fig_disp, use_container_width=True)

        with col_absrp:
            st.markdown("##### 📈 损失吸收效应最低资本占比")
            # 显示目标公司数值（箱型图上方）
            if target_co and "特定类别保险合同损失吸收效应" in df.columns and "最低资本" in df.columns:
                t_mask = df["公司"] == target_co
                if t_mask.any():
                    t_row = df[t_mask].iloc[0]
                    if pd.notna(t_row["最低资本"]) and t_row["最低资本"] != 0:
                        t_val = t_row["特定类别保险合同损失吸收效应"] / t_row["最低资本"]
                        if pd.notna(t_val):
                            st.info(f"🎯 **目标公司：{target_co}** — 损失吸收效应最低资本占比：**{t_val*100:.1f}%**")
            fig_absrp = boxplot_with_annotations(
                df, "特定类别保险合同损失吸收效应/最低资本",
                "损失吸收效应最低资本占比",
                height=360, target_co=target_co,
                group_by_category=True, y_min=-0.5, y_max=0.1, pct_display=True
            )
            if fig_absrp:
                st.plotly_chart(fig_absrp, use_container_width=True)

        st.markdown('<div class="section-break"></div>', unsafe_allow_html=True)

        # 保险风险最低资本占认可负债率
        st.markdown("#### 📊 保险风险最低资本占认可负债率")

        # 动态生成说明文字
        stress_desc = ""
        if "寿险业务保险风险最低资本合计/认可负债" in df.columns:
            stress_medians = {}
            for cat, cos in st.session_state.get("comp_cats", {}).items():
                vals = df[df["公司"].isin(cos)]["寿险业务保险风险最低资本合计/认可负债"].dropna()
                if len(vals) > 0:
                    stress_medians[cat] = float(vals.median())
            if stress_medians:
                sorted_stress = sorted(stress_medians.items(), key=lambda x: x[1])
                lowest = sorted_stress[0]
                low_cats = [cat for cat, v in sorted_stress if v <= lowest[1] + 0.01]
                stress_desc = "、".join(low_cats) + "公司的寿险保险风险压力因子明显低于其他类型公司的水平"

        desc_html = """<div class="metric-explain">
            • 以该指标体现公司保险风险最低资本相对认可负债的压力因子值，比率越高说明负债对应的保险风险水平越高；<br>
        """
        if stress_desc:
            desc_html += f"    • {stress_desc}。\n"
        desc_html += "</div>"
        st.markdown(desc_html, unsafe_allow_html=True)

        # 两列布局：左寿险，右非寿险
        col_life_stress, col_nonlife_stress = st.columns(2)

        with col_life_stress:
            st.markdown("##### 📈 保险风险（寿）/认可负债率")
            # 显示目标公司数值（箱型图上方）
            if target_co and "寿险业务保险风险最低资本合计" in df.columns and "认可负债" in df.columns:
                t_mask = df["公司"] == target_co
                if t_mask.any():
                    t_row = df[t_mask].iloc[0]
                    if pd.notna(t_row["认可负债"]) and t_row["认可负债"] != 0:
                        t_val = t_row["寿险业务保险风险最低资本合计"] / t_row["认可负债"]
                        if pd.notna(t_val):
                            st.info(f"🎯 **目标公司：{target_co}** — 保险风险（寿）/认可负债率：**{t_val*100:.1f}%**")
            fig_life_stress = boxplot_with_annotations(
                df, "寿险业务保险风险最低资本合计/认可负债",
                "保险风险（寿）/认可负债率",
                height=360, target_co=target_co,
                group_by_category=True, y_min=0, y_max=0.2, pct_display=True
            )
            if fig_life_stress:
                st.plotly_chart(fig_life_stress, use_container_width=True)

        with col_nonlife_stress:
            st.markdown("##### 📈 保险风险（非寿）/认可负债率")
            # 显示目标公司数值（箱型图上方）
            if target_co and "非寿险业务保险风险最低资本合计" in df.columns and "认可负债" in df.columns:
                t_mask = df["公司"] == target_co
                if t_mask.any():
                    t_row = df[t_mask].iloc[0]
                    if pd.notna(t_row["认可负债"]) and t_row["认可负债"] != 0:
                        t_val = t_row["非寿险业务保险风险最低资本合计"] / t_row["认可负债"]
                        if pd.notna(t_val):
                            st.info(f"🎯 **目标公司：{target_co}** — 保险风险（非寿）/认可负债率：**{t_val*100:.1f}%**")
            fig_nonlife_stress = boxplot_with_annotations(
                df, "非寿险业务保险风险最低资本合计/认可负债",
                "保险风险（非寿）/认可负债率",
                height=360, target_co=target_co,
                group_by_category=True, y_min=0, y_max=0.01, pct_display=True
            )
            if fig_nonlife_stress:
                st.plotly_chart(fig_nonlife_stress, use_container_width=True)

# —— 04-重大融资信息统计 ——

def render_page_04(standalone=True):
    """04-重大融资信息统计"""
    if not standalone:
        st.markdown('<div style="page-break-after: always;"></div>', unsafe_allow_html=True)
    st.markdown(f"### 📋 04 · 重大融资信息统计 · {q_label}")

    st.markdown("""
    <div class="metric-explain">
        <strong>说明：</strong>本页面统计各公司在各季度的增资/发债情况，以及综合偿付能力充足率的变动情况。
    </div>
    """, unsafe_allow_html=True)

    # 从Excel中读取重大融资信息（可选，文件不存在时跳过）
    # 优先从 session_state（上传的文件）读取，其次尝试本地文件
    df_fin = None
    
    if st.session_state.get("financing_file_bytes") is not None:
        try:
            import io as _io
            df_fin = pd.read_excel(_io.BytesIO(st.session_state.financing_file_bytes))
        except Exception as e:
            st.warning(f"⚠️ 读取重大融资信息失败: {e}")
    elif os.path.exists("重大融资信息.xlsx"):
        df_fin = pd.read_excel("重大融资信息.xlsx")
    
    if df_fin is not None and len(df_fin) > 0:
        
        # 解析"综合充足率变动"列，拆分为两列
        def parse_change(change_str):
            if pd.isna(change_str):
                return "", ""
            # Split on newline character
            parts = str(change_str).split("\n")
            total_change = ""
            financing_impact = ""
            for part in parts:
                if "季度总变动" in part:
                    total_change = part.replace("季度总变动：", "").strip()
                if "增资/发债的影响" in part:
                    financing_impact = part.replace("增资/发债的影响：", "").strip()
            return total_change, financing_impact

        df_fin[["季度总变动", "增资/发债的影响"]] = df_fin["综合充足率变动"].apply(
            lambda x: pd.Series(parse_change(x))
        )

        # 按季度倒序排列：解析季度字符串为排序键
        def parse_quarter(q_str):
            # "2025Q4" -> (2025, 4)
            year = int(q_str[:4])
            q = int(q_str[5:])
            return (year, q)

        df_fin["季度排序"] = df_fin["季度"].apply(parse_quarter)
        df_fin = df_fin.sort_values("季度排序", ascending=False).drop("季度排序", axis=1)

        # 统计摘要
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("融资事件总数", len(df_fin))
        with col2:
            st.metric("涉及公司数", df_fin["公司名称"].nunique())
        with col3:
            latest_q = df_fin["季度"].iloc[0] if len(df_fin) > 0 else "无"
            st.metric("最新季度", latest_q)

        st.divider()

        # 按季度分组展示
        st.markdown("##### 📊 重大融资信息明细")

        # 获取唯一的季度列表（已按倒序排列）
        quarters = df_fin["季度"].unique()

        for q in quarters:
            q_df = df_fin[df_fin["季度"] == q]

            with st.expander(f"📅 {q} （{len(q_df)} 家公司）", expanded=(q == quarters[0])):
                # 显示该季度的数据（使用HTML表格避免滚动条）
                display_df = q_df[["公司名称", "增资/发债", "季度总变动", "增资/发债的影响"]].copy()

                # 构建HTML表格（无滚动条，展示全部内容，使用内联样式）
                rows_html = []
                for _, row in display_df.iterrows():
                    rows_html.append(
                        f'<tr style="border-bottom:1px solid #eee;">'
                        f'<td style="padding:8px 12px; color:#333;">{row["公司名称"]}</td>'
                        f'<td style="padding:8px 12px; color:#333;">{row["增资/发债"]}</td>'
                        f'<td style="padding:8px 12px; color:#333;">{row["季度总变动"]}</td>'
                        f'<td style="padding:8px 12px; color:#333;">{row["增资/发债的影响"]}</td>'
                        f'</tr>'
                    )

                table_html = (
                    f'<table style="width:100%; border-collapse:collapse; font-size:0.9rem; margin:8px 0;">'
                    f'<thead>'
                    f'<tr style="background:#1A3A5C; color:white;">'
                    f'<th style="padding:10px 12px; text-align:left; font-weight:600;">公司名称</th>'
                    f'<th style="padding:10px 12px; text-align:left; font-weight:600;">增资/发债</th>'
                    f'<th style="padding:10px 12px; text-align:left; font-weight:600;">季度总变动</th>'
                    f'<th style="padding:10px 12px; text-align:left; font-weight:600;">增资/发债的影响</th>'
                    f'</tr>'
                    f'</thead>'
                    f'<tbody>{"".join(rows_html)}</tbody>'
                    f'</table>'
                )
                st.markdown(table_html, unsafe_allow_html=True)

        # 下载按钮
        st.divider()
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df_fin.to_excel(w, index=False, sheet_name="重大融资信息")
        buf.seek(0)
        st.download_button(
            "⬇ 下载重大融资信息（Excel）",
            data=buf,
            file_name=f"重大融资信息_{q_label}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    else:
        st.info("📁 未找到重大融资信息数据，跳过此部分。")
        st.caption("💡 如需显示重大融资信息，请在左侧边栏「重大融资信息」处上传 Excel 文件。")






def boxplot_with_annotations(df, indicator, yaxis_title, height=360, target_co=None, y_multiplier=1.0, group_by_category=False, y_min=0, y_max=None, pct_display=False, dtick=None):
    """
    画箱型图（完全自定义绘制，参照参考图风格）
    - 使用 add_shape + add_annotation 手动绘制
    - 包含 whisker 虚线 + 末端小横线（参照参考图）
    - 空心箱体、中位线、min/max 公司名称标注
    """
    import plotly.graph_objects as go

    sr = df[indicator].dropna()
    if sr.empty:
        return None

    vals = sr * y_multiplier
    co_names = df.loc[sr.index, "公司"].reset_index(drop=True)
    vals = vals.reset_index(drop=True)

    # 分类颜色（参照参考图）
    CAT_COLORS = {
        "大型公司": "#2E7AD6",
        "中型公司": "#7B5DBE",
        "小型公司": "#E8913A",
        "银行系":   "#2EB872",
        "外资系":   "#E24B4A",
        "养老健康": "#34495E",
    }

    fig = go.Figure()

    if group_by_category:
        # ========== 按公司分类分组绘制 ==========
        # 固定分类顺序：大型公司、中型公司、小型公司、银行系、外资系、养老健康
        fixed_cat_order = ["大型公司", "中型公司", "小型公司", "银行系", "外资系", "养老健康"]
        cat_names_ordered = []
        stats = {}
        
        for cat_name in fixed_cat_order:
            if cat_name not in st.session_state.get("comp_cats", {}):
                continue
            cat_cos = st.session_state.get("comp_cats", {}).get(cat_name, [])
            mask = co_names.isin(cat_cos)
            cat_vals = vals[mask]
            if cat_vals.empty:
                continue
            
            cat_names_ordered.append(cat_name)
            color = CAT_COLORS.get(cat_name, "#2E7AD6")

            q1_v = float(cat_vals.quantile(0.25))
            med_v = float(cat_vals.median())
            q3_v = float(cat_vals.quantile(0.75))

            max_idx = int(cat_vals.idxmax())
            min_idx = int(cat_vals.idxmin())
            max_co = str(co_names.iloc[max_idx])
            min_co = str(co_names.iloc[min_idx])
            max_val = float(cat_vals.max())
            min_val = float(cat_vals.min())

            stats[cat_name] = {
                'q1': q1_v, 'median': med_v, 'q3': q3_v,
                'color': color,
                'max_co': max_co, 'min_co': min_co,
                'max_val': max_val, 'min_val': min_val,
            }

        n_cats = len(cat_names_ordered)
        if n_cats == 0:
            return None

        BOX_W = 0.50
        TICK_LEN = 0.08  # whisker 末端小横线半长

        for i, cat_name in enumerate(cat_names_ordered):
            s = stats[cat_name]
            x_center = i
            c = s['color']
            x0 = x_center - BOX_W / 2
            x1 = x_center + BOX_W / 2
            q1 = s['q1']
            med = s['median']
            q3 = s['q3']
            mn = s['min_val']
            mx = s['max_val']

            # ---- 1) 下 whisker（虚线，加粗至1.2px）----
            fig.add_shape(
                type="line",
                x0=x_center, x1=x_center, y0=mn, y1=q1,
                line=dict(color=c, width=1.2, dash="dot"),
                layer="above",
            )
            # 下 whisker 末端小横线（加粗至1.8px）
            fig.add_shape(
                type="line",
                x0=x_center - TICK_LEN, x1=x_center + TICK_LEN, y0=mn, y1=mn,
                line=dict(color=c, width=1.8),
                layer="above",
            )

            # ---- 2) 上 whisker（虚线，加粗至1.2px）----
            fig.add_shape(
                type="line",
                x0=x_center, x1=x_center, y0=q3, y1=mx,
                line=dict(color=c, width=1.2, dash="dot"),
                layer="above",
            )
            # 上 whisker 末端小横线（加粗至1.8px）
            fig.add_shape(
                type="line",
                x0=x_center - TICK_LEN, x1=x_center + TICK_LEN, y0=mx, y1=mx,
                line=dict(color=c, width=1.8),
                layer="above",
            )

            # ---- 3) 箱体（极浅蓝色调填充）----
            # 使用统一的极浅填充色，保留分类边框颜色
            fig.add_shape(
                type="rect",
                x0=x0, x1=x1, y0=q1, y1=q3,
                fillcolor="rgba(245, 248, 255, 0.55)",
                line=dict(color=c, width=1.5),
                layer="above",
            )

            # ---- 4) 中位线（水平实线，加粗至2.5px）----
            fig.add_shape(
                type="line",
                x0=x0, x1=x1, y0=med, y1=med,
                line=dict(color=c, width=2.5),
                layer="above",
            )

            # ---- 5) 数值标注（箱体内部）已移除 ----

            # ---- 6) min/max 公司名称标注（whisker 末端）----
            # 最大值（上 whisker 末端）— 如果超过 y_max 则在边界处显示
            if mx > q3:
                display_mx = mx if (y_max is None or mx <= y_max) else y_max
                # 标注显示真实值，位置在边界处
                lbl_mx = f"{mx*100:.1f}%" if pct_display else f"{mx:.1f}%"
                yshift_mx = 10 if (y_max is None or mx <= y_max) else 5
                fig.add_annotation(
                    x=x_center, y=display_mx,
                    text=f"{s['max_co']} {lbl_mx}",
                    showarrow=False,
                    yshift=yshift_mx,
                    font=dict(size=10, color=c, family="SimHei"),
                )

            # 最小值（下 whisker 末端）— 如果小于 y_min 则在边界处显示
            if mn < q1:
                display_mn = mn if (mn >= y_min) else y_min
                # 标注显示真实值，位置在边界处
                lbl_mn = f"{mn*100:.1f}%" if pct_display else f"{mn:.1f}%"
                yshift_mn = -10 if (mn >= y_min) else -5
                fig.add_annotation(
                    x=x_center, y=display_mn,
                    text=f"{s['min_co']} {lbl_mn}",
                    showarrow=False,
                    yshift=yshift_mn,
                    font=dict(size=10, color=c, family="SimHei"),
                )

            # ---- 7) 图例 trace ----
            fig.add_trace(go.Scatter(
                x=[None], y=[None],
                mode="markers",
                marker=dict(size=10, color=c, symbol="square"),
                name=cat_name,
                showlegend=True,
            ))

        # ---------- 目标公司高亮（圆点 + 文字，参照参考图） ----------
        if target_co and len(target_co) > 0:
            t_mask = (co_names == target_co)
            if t_mask.any():
                t_val = float(vals[t_mask].iloc[0])
                t_cat = None
                for cat_name, cat_cos in st.session_state.get("comp_cats", {}).items():
                    if target_co in cat_cos:
                        t_cat = cat_name
                        break
                if t_cat and t_cat in cat_names_ordered:
                    t_idx = cat_names_ordered.index(t_cat)
                    t_color = stats[t_cat]['color']
                    display_y = min(t_val, y_max) if y_max is not None else t_val
                    lbl_t = f"{t_val*100:.1f}%" if pct_display else f"{t_val:.1f}%"
                    fig.add_trace(go.Scatter(
                        x=[t_idx],
                        y=[display_y],
                        mode="markers+text",
                        marker=dict(size=16, color="#00BFFF", line=dict(color="#fff", width=1.5), symbol="star"),
                        text=[f"{target_co} {lbl_t}"],
                        textposition="top center",
                        textfont=dict(size=11, color="#00BFFF", family="SimHei"),
                        showlegend=False,
                        hoverinfo="skip",
                    ))

        # ---------- 布局 ----------
        yaxis_cfg = dict(
            gridcolor="#f5f5f5",
            zeroline=False,
            title_font=dict(size=13, family="SimHei"),
            tickfont=dict(size=10, family="SimHei"),
            tickformat=".0%" if pct_display or (y_multiplier == 1.0 and vals.max() < 2) else "",
        )
        if y_max is not None:
            yaxis_cfg["range"] = [y_min, y_max]
            # 显式设置 dtick，防止小范围百分比刻度出现重复标签
            if dtick is not None:
                yaxis_cfg["dtick"] = dtick
            elif pct_display:
                y_span = y_max - y_min
                if y_span <= 0.05:
                    yaxis_cfg["dtick"] = 0.01
                elif y_span <= 0.15:
                    yaxis_cfg["dtick"] = 0.02
                elif y_span <= 0.5:
                    yaxis_cfg["dtick"] = 0.05
                elif y_span <= 1.0:
                    yaxis_cfg["dtick"] = 0.1
                elif y_span <= 5.0:
                    yaxis_cfg["dtick"] = 0.5
                else:
                    yaxis_cfg["dtick"] = 2.0

        fig.update_layout(
            height=height,
            margin=dict(l=0, r=120, t=0, b=0),
            plot_bgcolor="white",
            paper_bgcolor="white",
            font=dict(family="SimHei, Microsoft YaHei, sans-serif"),
            yaxis_title=yaxis_title,
            xaxis=dict(
                tickmode="array",
                tickvals=list(range(n_cats)),
                ticktext=cat_names_ordered,
                tickfont=dict(size=10, family="SimHei"),
                showgrid=False,
                range=[-0.5, n_cats - 0.6],
            ),
            yaxis=yaxis_cfg,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5,
                font=dict(size=12, family="SimHei"),
            ),
            showlegend=True,
            hovermode=False,
        )

    else:
        # ========== 单箱型图 ==========
        q1_v = float(vals.quantile(0.25))
        med_v = float(vals.median())
        q3_v = float(vals.quantile(0.75))
        mn = float(vals.min())
        mx = float(vals.max())
        color = "#2E7AD6"
        BOX_W = 0.50
        TICK_LEN = 0.08
        x_center = 0
        x0 = x_center - BOX_W / 2
        x1 = x_center + BOX_W / 2

        # 下 whisker
        fig.add_shape(type="line", x0=x_center, x1=x_center, y0=mn, y1=q1_v,
                      line=dict(color=color, width=1, dash="dot"), layer="above")
        fig.add_shape(type="line", x0=x_center-TICK_LEN, x1=x_center+TICK_LEN, y0=mn, y1=mn,
                      line=dict(color=color, width=1.5), layer="above")
        # 上 whisker
        fig.add_shape(type="line", x0=x_center, x1=x_center, y0=q3_v, y1=mx,
                      line=dict(color=color, width=1, dash="dot"), layer="above")
        fig.add_shape(type="line", x0=x_center-TICK_LEN, x1=x_center+TICK_LEN, y0=mx, y1=mx,
                      line=dict(color=color, width=1.5), layer="above")
        # 箱体
        fig.add_shape(type="rect", x0=x0, x1=x1, y0=q1_v, y1=q3_v,
                      fillcolor="rgba(255,255,255,0.3)", line=dict(color=color, width=1.5), layer="above")
        # 中位线
        fig.add_shape(type="line", x0=x0, x1=x1, y0=med_v, y1=med_v,
                      line=dict(color=color, width=2), layer="above")

        # 标注已移除（不显示箱体内部数值）

        # min/max 标注（超出Y轴范围时在边界处显示真实值）
        if mx > q3_v:
            display_mx = mx if (y_max is None or mx <= y_max) else y_max
            # 标注显示真实值，位置在边界处
            lbl_mx = f"{mx*100:.1f}%" if pct_display else f"{mx:.1f}%"
            yshift_mx = 10 if (y_max is None or mx <= y_max) else 5
            fig.add_annotation(x=x_center, y=display_mx, text=f"{co_names.iloc[int(vals.idxmax())]} {lbl_mx}",
                               showarrow=False, yshift=yshift_mx, font=dict(size=10, color=color, family="SimHei"))
        if mn < q1_v:
            display_mn = mn if (mn >= y_min) else y_min
            # 标注显示真实值，位置在边界处
            lbl_mn = f"{mn*100:.1f}%" if pct_display else f"{mn:.1f}%"
            yshift_mn = -10 if (mn >= y_min) else -5
            fig.add_annotation(x=x_center, y=display_mn, text=f"{co_names.iloc[int(vals.idxmin())]} {lbl_mn}",
                               showarrow=False, yshift=yshift_mn, font=dict(size=10, color=color, family="SimHei"))

        # 目标公司高亮（超出范围时在边界处显示）
        if target_co and len(target_co) > 0:
            t_mask = (co_names == target_co)
            if t_mask.any():
                t_val = float(vals[t_mask].iloc[0])
                display_y = t_val if (y_max is None or t_val <= y_max) else y_max
                display_y = display_y if display_y >= y_min else y_min
                lbl_t = f"{t_val*100:.1f}%" if pct_display else f"{t_val:.1f}%"
                fig.add_trace(go.Scatter(
                    x=[x_center], y=[display_y],
                    mode="markers+text",
                    marker=dict(size=16, color="#185FA5", line=dict(color="#fff", width=1.5), symbol="star"),
                    text=[f"{target_co} {lbl_t}"],
                    textposition="top center",
                    textfont=dict(size=11, color="#185FA5", family="SimHei"),
                    showlegend=False, hoverinfo="skip",
                ))

        yaxis_cfg = dict(
            gridcolor="#e0e0e0",
            zeroline=False,
            title_font=dict(size=13, family="SimHei"),
            tickformat=".0%" if pct_display else "",
        )
        if y_max is not None:
            yaxis_cfg["range"] = [y_min, y_max]
            # 显式设置 dtick，防止小范围百分比刻度出现重复标签
            if dtick is not None:
                yaxis_cfg["dtick"] = dtick
            elif pct_display:
                y_span = y_max - y_min
                if y_span <= 0.05:
                    yaxis_cfg["dtick"] = 0.01
                elif y_span <= 0.15:
                    yaxis_cfg["dtick"] = 0.02
                elif y_span <= 0.5:
                    yaxis_cfg["dtick"] = 0.05
                elif y_span <= 1.0:
                    yaxis_cfg["dtick"] = 0.1
                elif y_span <= 5.0:
                    yaxis_cfg["dtick"] = 0.5
                else:
                    yaxis_cfg["dtick"] = 2.0

        fig.update_layout(
            height=height,
            margin=dict(l=0, r=120, t=0, b=0),
            plot_bgcolor="white",
            paper_bgcolor="white",
            font=dict(family="SimHei, Microsoft YaHei, sans-serif"),
            yaxis_title=yaxis_title,
            xaxis_visible=False,
            yaxis=yaxis_cfg,
            showlegend=False,
            hovermode=False,
        )

    return fig


def capital_tier_boxplot(df, target_co=None, height=360):
    """
    资本分级行业分布箱线图（02页面专用）
    - 展示核心一级/二级、附属一级/二级资本占实际资本的比例分布
    - 4个箱体并排，使用不同颜色区分
    - 风格与 boxplot_with_annotations 保持一致
    """
    import plotly.graph_objects as go

    # 计算各分级资本占实际资本的比例
    df = df.copy()
    df["核心一级资本占比"] = df["核心一级资本"] / df["实际资本"]
    df["核心二级资本占比"] = df["核心二级资本"] / df["实际资本"]
    df["附属一级资本占比"] = df["附属一级资本"] / df["实际资本"]
    df["附属二级资本占比"] = df["附属二级资本"] / df["实际资本"]

    tiers = [
        ("核心一级资本", "#2E7AD6", "核心一级资本占比"),
        ("核心二级资本", "#7B5DBE", "核心二级资本占比"),
        ("附属一级资本", "#E8913A", "附属一级资本占比"),
        ("附属二级资本", "#2EB872", "附属二级资本占比"),
    ]

    fig = go.Figure()
    BOX_W = 0.35
    TICK_LEN = 0.08

    for i, (label, color, col) in enumerate(tiers):
        sr = df[col].dropna()
        if sr.empty:
            continue

        q1 = float(sr.quantile(0.25))
        med = float(sr.median())
        q3 = float(sr.quantile(0.75))
        mn = float(sr.min())
        mx = float(sr.max())

        x_center = i
        x0 = x_center - BOX_W / 2
        x1 = x_center + BOX_W / 2

        # 下 whisker（虚线，加粗至1.2px）
        fig.add_shape(type="line", x0=x_center, x1=x_center, y0=mn, y1=q1,
                      line=dict(color=color, width=1.2, dash="dot"), layer="above")
        fig.add_shape(type="line", x0=x_center-TICK_LEN, x1=x_center+TICK_LEN, y0=mn, y1=mn,
                      line=dict(color=color, width=1.8), layer="above")
        # 上 whisker（虚线，加粗至1.2px）
        fig.add_shape(type="line", x0=x_center, x1=x_center, y0=q3, y1=mx,
                      line=dict(color=color, width=1.2, dash="dot"), layer="above")
        fig.add_shape(type="line", x0=x_center-TICK_LEN, x1=x_center+TICK_LEN, y0=mx, y1=mx,
                      line=dict(color=color, width=1.8), layer="above")
        # 箱体（极浅蓝色调填充）
        fig.add_shape(type="rect", x0=x0, x1=x1, y0=q1, y1=q3,
                      fillcolor="rgba(245, 248, 255, 0.55)", line=dict(color=color, width=1.5), layer="above")
        # 中位线（加粗至2.5px）
        fig.add_shape(type="line", x0=x0, x1=x1, y0=med, y1=med,
                      line=dict(color=color, width=2.5), layer="above")

        # 中位数标注（箱体内部）
        fig.add_annotation(x=x_center, y=med, text=f"{med*100:.0f}%",
                           showarrow=False, font=dict(size=10, color=color, family="SimHei"))

        # 上四分位标注
        fig.add_annotation(x=x_center, y=q3, text=f"{q3*100:.0f}%",
                           showarrow=False, yshift=8, font=dict(size=10, color=color, family="SimHei"))

        # 下四分位标注
        fig.add_annotation(x=x_center, y=q1, text=f"{q1*100:.0f}%",
                           showarrow=False, yshift=-8, font=dict(size=10, color=color, family="SimHei"))

        # max 标注（如果超过Q3）
        if mx > q3:
            max_idx = int(sr.idxmax())
            max_co = str(df.loc[max_idx, "公司"])
            fig.add_annotation(x=x_center, y=mx, text=f"{max_co} {mx*100:.0f}%",
                               showarrow=False, yshift=10, font=dict(size=10, color=color, family="SimHei"))

        # min 标注（如果小于Q1且>=0）
        if mn < q1 and mn >= 0:
            min_idx = int(sr.idxmin())
            min_co = str(df.loc[min_idx, "公司"])
            fig.add_annotation(x=x_center, y=mn, text=f"{min_co} {mn*100:.0f}%",
                               showarrow=False, yshift=-10, font=dict(size=10, color=color, family="SimHei"))

        # 图例 trace
        fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers",
                                 marker=dict(size=10, color=color, symbol="square"),
                                 name=label, showlegend=True))

        # 目标公司高亮
        if target_co and len(target_co) > 0:
            t_mask = (df["公司"] == target_co)
            if t_mask.any():
                t_val = float(df.loc[t_mask, col].iloc[0])
                fig.add_trace(go.Scatter(
                    x=[x_center], y=[t_val],
                    mode="markers+text",
                    marker=dict(size=16, color="#00BFFF", line=dict(color="#fff", width=1.5), symbol="star"),
                    text=[f"{target_co} {t_val*100:.0f}%"],
                    textposition="top center",
                    textfont=dict(size=10, color="#00BFFF", family="SimHei"),  # 目标公司标签10px
                    showlegend=False, hoverinfo="skip",
                ))

    fig.update_layout(
        height=height,
        margin=dict(l=0, r=80, t=0, b=0),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="SimHei, Microsoft YaHei, sans-serif"),
        yaxis_title="占实际资本比例",
        yaxis=dict(gridcolor="#f5f5f5", zeroline=False, range=[-0.02, 1.05],
                   tickformat=".0%", title_font=dict(size=13, family="SimHei"),  # Y轴标题13px
                   tickfont=dict(size=10, family="SimHei")),  # Y轴刻度10px
        xaxis=dict(tickmode="array", tickvals=list(range(len(tiers))),
                   ticktext=[t[0] for t in tiers], tickfont=dict(size=10, family="SimHei"),
                   showgrid=False, range=[-0.5, len(tiers)-0.8]),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.08,
            xanchor="center",
            x=0.5,
            font=dict(size=11, family="SimHei"),
        ),
        showlegend=True, hovermode=False,
    )
    return fig

def mc_composition_boxplot(df, indicators, denominator_col, target_co=None, height=360):
    """
    最低资本构成分布箱型图（03页面专用）
    - 固定为全行业汇总展示，不随公司类型筛选变化
    - 手动绘制风格（与01/02页面一致）：空心箱体、虚线whisker、末端小横线
    - Y轴范围固定为[-60%, 100%]
    """
    import plotly.graph_objects as go
    import pandas as pd

    n_inds = len(indicators)
    if n_inds == 0 or denominator_col not in df.columns:
        return None

    fig = go.Figure()

    # 固定Y轴范围
    Y_MIN = -0.6
    Y_MAX = 1.0

    BOX_W = 0.35
    TICK_LEN = 0.08

    # 计算每个指标的统计量
    stats_list = []
    for j, ind in enumerate(indicators):
        if ind["col"] not in df.columns:
            continue
        col_vals = df[ind["col"]] / df[denominator_col]
        valid = col_vals.dropna()
        if len(valid) == 0:
            continue

        q1 = float(valid.quantile(0.25))
        med = float(valid.quantile(0.50))
        q3 = float(valid.quantile(0.75))
        mn = float(valid.min())
        mx = float(valid.max())

        max_idx = int(valid.idxmax())
        min_idx = int(valid.idxmin())
        max_co = str(df.loc[max_idx, "公司"])
        min_co = str(df.loc[min_idx, "公司"])

        stats_list.append({
            "label": ind["label"],
            "color": ind.get("color", "#2E7AD6"),
            "q1": q1, "median": med, "q3": q3,
            "min": mn, "max": mx,
            "min_co": min_co, "max_co": max_co,
            "col": ind["col"],
        })

    n_stats = len(stats_list)
    if n_stats == 0:
        return None

    for i, s in enumerate(stats_list):
        x_center = i
        c = s["color"]
        x0 = x_center - BOX_W / 2
        x1 = x_center + BOX_W / 2
        q1 = s["q1"]
        med = s["median"]
        q3 = s["q3"]
        mn = s["min"]
        mx = s["max"]

        # ---- 1) 下 whisker（虚线）----
        fig.add_shape(
            type="line",
            x0=x_center, x1=x_center, y0=mn, y1=q1,
            line=dict(color=c, width=1, dash="dot"),
            layer="above",
        )
        # 下 whisker 末端小横线
        fig.add_shape(
            type="line",
            x0=x_center - TICK_LEN, x1=x_center + TICK_LEN, y0=mn, y1=mn,
            line=dict(color=c, width=1.5),
            layer="above",
        )

        # ---- 2) 上 whisker（虚线）----
        fig.add_shape(
            type="line",
            x0=x_center, x1=x_center, y0=q3, y1=mx,
            line=dict(color=c, width=1, dash="dot"),
            layer="above",
        )
        # 上 whisker 末端小横线
        fig.add_shape(
            type="line",
            x0=x_center - TICK_LEN, x1=x_center + TICK_LEN, y0=mx, y1=mx,
            line=dict(color=c, width=1.5),
            layer="above",
        )

        # ---- 3) 箱体（空心矩形）----
        fig.add_shape(
            type="rect",
            x0=x0, x1=x1, y0=q1, y1=q3,
            fillcolor="rgba(255,255,255,0.3)",
            line=dict(color=c, width=1.5),
            layer="above",
        )

        # ---- 4) 中位线（水平实线）----
        fig.add_shape(
            type="line",
            x0=x0, x1=x1, y0=med, y1=med,
            line=dict(color=c, width=2),
            layer="above",
        )

        # ---- 5) max 标注（上 whisker 末端）----
        display_mx = mx if mx <= Y_MAX else Y_MAX
        yshift_mx = 10 if mx <= Y_MAX else 5
        fig.add_annotation(
            x=x_center, y=display_mx,
            text=f"{s['max_co']} {mx*100:.1f}%",
            showarrow=False,
            yshift=yshift_mx,
            font=dict(size=10, color=c, family="SimHei"),
        )

        # ---- 6) min 标注（下 whisker 末端）----
        display_mn = mn if mn >= Y_MIN else Y_MIN
        yshift_mn = -10 if mn >= Y_MIN else -5
        fig.add_annotation(
            x=x_center, y=display_mn,
            text=f"{s['min_co']} {mn*100:.1f}%",
            showarrow=False,
            yshift=yshift_mn,
            font=dict(size=10, color=c, family="SimHei"),
        )

        # ---- 7) 图例 trace ----
        fig.add_trace(go.Scatter(
            x=[None], y=[None],
            mode="markers",
            marker=dict(size=10, color=c, symbol="square"),
            name=s["label"],
            showlegend=True,
        ))

    # ---- 目标公司高亮 ----
    if target_co and len(target_co) > 0:
        t_mask = (df["公司"] == target_co)
        if t_mask.any():
            trow = df[t_mask].iloc[0]
            if denominator_col in trow.index:
                denom = trow[denominator_col]
                if pd.notna(denom) and denom != 0:
                    for i, s in enumerate(stats_list):
                        if s["col"] in trow.index:
                            t_val = trow[s["col"]] / denom
                            if pd.notna(t_val):
                                display_y = t_val if t_val <= Y_MAX else Y_MAX
                                display_y = display_y if display_y >= Y_MIN else Y_MIN
                                fig.add_trace(go.Scatter(
                                    x=[i],
                                    y=[display_y],
                                    mode="markers+text",
                                    marker=dict(size=16, color="#00BFFF", line=dict(color="#fff", width=1.5), symbol="star"),
                                    text=[f"{target_co} {t_val*100:.1f}%"],
                                    textposition="top center",
                                    textfont=dict(size=10, color="#00BFFF", family="SimHei"),
                                    showlegend=False,
                                    hoverinfo="skip",
                                ))

    fig.update_layout(
        height=height,
        margin=dict(l=0, r=80, t=0, b=0),
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(family="SimHei, Microsoft YaHei, sans-serif"),
        yaxis=dict(
            tickformat=".0%",
            gridcolor="#e0e0e0",
            zeroline=False,
            title_font=dict(size=13, family="SimHei"),
            tickfont=dict(size=10, family="SimHei"),  # Y轴刻度10px
            range=[Y_MIN, Y_MAX],
        ),
        xaxis=dict(
            tickmode="array",
            tickvals=list(range(n_stats)),
            ticktext=[s["label"] for s in stats_list],
            tickfont=dict(size=10, family="SimHei"),
            showgrid=False,
            range=[-0.5, n_stats - 0.8],
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.08,
            xanchor="center",
            x=0.5,
            font=dict(size=11, family="SimHei"),
        ),
        showlegend=True,
        hovermode=False,
    )

    return fig


def histogram_with_box(df, indicator, xaxis_title, yaxis_title="公司数", nbins=25, target_co=None, val_multiplier=1.0):
    """分布直方图 + 上方箱型图（plotly express版，支持目标公司高亮）"""
    import plotly.express as px
    import pandas as pd

    sr = df[indicator].dropna()
    if sr.empty:
        return None

    col_plot = "__plot_val__"
    dfp = pd.DataFrame({
        "__plot_val__": sr * val_multiplier,
        "公司": df.loc[sr.index, "公司"].values
    })

    fig = px.histogram(
        dfp, x=col_plot, nbins=nbins,
        color_discrete_sequence=["#185FA5"],
        marginal="box",
        hover_data=["公司"]
    )

    # 目标公司高亮：在直方图上方标注
    if target_co and target_co in dfp["公司"].values:
        t_row = dfp[dfp["公司"] == target_co]
        if not t_row.empty:
            t_val = t_row.iloc[0][col_plot]
            fig.add_vline(
                x=t_val,
                line_color="#EF9F27",
                line_width=2,
                annotation_text=f"  {target_co}",
                annotation_position="top"
            )

    fig.update_layout(
        height=360,
        margin=dict(l=0, r=60, t=30, b=0),
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis_title=xaxis_title,
        yaxis_title=yaxis_title,
        yaxis=dict(gridcolor="#e8e8e8"),
        xaxis=dict(gridcolor="#e8e8e8"),
        showlegend=False,
    )
    return fig

def color_cells(v, kind="comp"):
    try:
        v = float(v)
    except:
        return "#888"
    if kind == "comp":
        if v < THR_COMP_PASS:  return "#E24B4A"
        if v < THR_COMP_WARN: return "#EF9F27"
        return "#1D9E75"
    return "#185FA5"

def risk_info(v):
    try:
        v = float(v)
    except:
        return "未知", "#888"
    if v >= THR_COMP_WARN: return "充足",   "#1D9E75"
    if v >= THR_COMP_PASS: return "预警", "#EF9F27"
    return "不达标", "#E24B4A"

# ===========================================================
# 登录权限体系
# ===========================================================
def init_login_state():
    if "logged_in"  not in st.session_state:
        st.session_state.logged_in  = False
    if "username"   not in st.session_state:
        st.session_state.username   = ""
    if "user_role"  not in st.session_state:
        st.session_state.user_role  = None
    if "ai_api_key" not in st.session_state:
        st.session_state.ai_api_key = ""
    if "ai_model"   not in st.session_state:
        st.session_state.ai_model   = "qwen-plus"

init_login_state()

if not st.session_state.logged_in:
    # 登录页专用样式：覆盖表单标签、输入框、按钮为浅色，并设置深蓝背景
    st.markdown("""
    <style>
    /* 登录页背景 */
    .stApp { background: #050b14 !important; }
    /* 登录页表单标签 */
    [data-testid="stForm"] label,
    [data-testid="stTextInput"] label,
    [data-testid="stTextInput"] > div > label,
    [data-baseweb="input"] label {
        color: #e0e0e0 !important;
    }
    /* 登录页输入框 */
    [data-testid="stTextInput"] input,
    [data-baseweb="input"] input {
        color: #333333 !important;
    }
    /* 登录页按钮 */
    [data-testid="stForm"] button[kind="formSubmit"] {
        background: linear-gradient(135deg,#0d5fa5,#00b4d8) !important;
        color: white !important;
        border: none !important;
    }
    </style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="login-container">', unsafe_allow_html=True)
    st.markdown('<p class="login-title">寿研数智<br>偿付能力季报分享平台</p>', unsafe_allow_html=True)
    st.markdown('<p class="login-sub">ACTUARIAL INTELLIGENCE</p>', unsafe_allow_html=True)

    # 使用三列布局缩窄表单宽度
    left_spacer, form_col, right_spacer = st.columns([1, 2, 1])
    with form_col:
        with st.form("login_form"):
            username = st.text_input("用户名", value="admin", key="login_user")
            password = st.text_input("密码", type="password", value="", key="login_pwd")
            submitted = st.form_submit_button("登录", use_container_width=True)

            if submitted:
                if password != LOGIN_PASSWORD:
                    st.error("密码错误，请联系管理员。")
                else:
                    st.session_state.logged_in  = True
                    st.session_state.username   = username
                    st.session_state.user_role   = "admin" if username in ADMIN_USERS else "user"
                    st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ===========================================================
# 数据上传与加载
# ===========================================================
def get_available_quarters(uploaded_file_bytes):
    """从上传的Excel字节流中获取所有符合季度格式的Sheet名称（如 2025Q4、2024Q3）"""
    import re
    try:
        import io
        xl_file = pd.ExcelFile(io.BytesIO(uploaded_file_bytes))
        # 只返回匹配 YYYYQN 格式的Sheet名称
        quarter_pattern = re.compile(r'^\d{4}Q[1-4]$')
        quarters = [s for s in xl_file.sheet_names if quarter_pattern.match(str(s).strip())]
        return sorted(quarters, reverse=True)  # 按时间倒序排列（最新在前）
    except Exception as e:
        return []

@st.cache_data
def get_df(uploaded_file_bytes, sheet_name):
    """从上传的Excel字节流中读取数据"""
    import io
    xl_file = io.BytesIO(uploaded_file_bytes)
    df = pd.read_excel(xl_file, sheet_name=sheet_name, header=0)
    df = clean_df(df)
    return df

# 初始化上传状态
if "uploaded_file_bytes" not in st.session_state:
    st.session_state.uploaded_file_bytes = None
if "uploaded_file_name" not in st.session_state:
    st.session_state.uploaded_file_name = None
if "available_quarters" not in st.session_state:
    st.session_state.available_quarters = []
# 重大融资信息文件上传状态初始化
if "financing_file_bytes" not in st.session_state:
    st.session_state.financing_file_bytes = None
if "financing_file_name" not in st.session_state:
    st.session_state.financing_file_name = None

# 检查是否已上传数据
if st.session_state.uploaded_file_bytes is not None:
    df_all = get_df(st.session_state.uploaded_file_bytes, st.session_state.get("current_quarter", "2025Q4"))
    
    # 动态读取公司分类（从当前季度的数据中）
    st.session_state.comp_cats = load_comp_cats_from_df(df_all)
    
    # 列存在性标记（供页面渲染使用）
    has_c = RATIO_COMP in df_all.columns
    has_k = RATIO_CORE in df_all.columns
    
    # 添加百分比辅助列
    df_all = df_all.copy()
    for src, dst in [(RATIO_COMP, COMP_PCT), (RATIO_CORE, CORE_PCT)]:
        if src in df_all.columns:
            df_all[dst] = df_all[src] * 100
    
    # 添加保险风险最低资本占比辅助列
    if "寿险业务保险风险最低资本合计" in df_all.columns and "最低资本" in df_all.columns:
        df_all["寿险业务保险风险最低资本合计/最低资本"] = df_all["寿险业务保险风险最低资本合计"] / df_all["最低资本"]
    if "非寿险业务保险风险最低资本合计" in df_all.columns and "最低资本" in df_all.columns:
        df_all["非寿险业务保险风险最低资本合计/最低资本"] = df_all["非寿险业务保险风险最低资本合计"] / df_all["最低资本"]
    if "市场风险-最低资本合计" in df_all.columns and "最低资本" in df_all.columns:
        df_all["市场风险-最低资本合计/最低资本"] = df_all["市场风险-最低资本合计"] / df_all["最低资本"]
    if "信用风险-最低资本合计" in df_all.columns and "最低资本" in df_all.columns:
        df_all["信用风险-最低资本合计/最低资本"] = df_all["信用风险-最低资本合计"] / df_all["最低资本"]
    if "市场风险-利率风险最低资本" in df_all.columns and "认可资产" in df_all.columns:
        df_all["利率风险/认可资产"] = df_all["市场风险-利率风险最低资本"] / df_all["认可资产"]
    if "市场风险-权益价格风险最低资本" in df_all.columns and "认可资产" in df_all.columns:
        df_all["权益价格风险/认可资产"] = df_all["市场风险-权益价格风险最低资本"] / df_all["认可资产"]
    if "信用风险-利差风险最低资本" in df_all.columns and "认可资产" in df_all.columns:
        df_all["利差风险/认可资产"] = df_all["信用风险-利差风险最低资本"] / df_all["认可资产"]
    if "信用风险-交易对手违约风险最低资本" in df_all.columns and "认可资产" in df_all.columns:
        df_all["对手违约风险/认可资产"] = df_all["信用风险-交易对手违约风险最低资本"] / df_all["认可资产"]
    if "量化风险分散效应" in df_all.columns and "最低资本" in df_all.columns:
        df_all["量化风险分散效应/最低资本"] = df_all["量化风险分散效应"] / df_all["最低资本"]
    if "特定类别保险合同损失吸收效应" in df_all.columns and "最低资本" in df_all.columns:
        df_all["特定类别保险合同损失吸收效应/最低资本"] = df_all["特定类别保险合同损失吸收效应"] / df_all["最低资本"]
else:
    df_all = None
    has_c = False
    has_k = False
    st.session_state.comp_cats = {}

# 添加更多辅助列（仅在数据已加载时执行）
if df_all is not None:
    if "寿险业务保险风险最低资本合计" in df_all.columns and "认可负债" in df_all.columns:
        df_all["寿险业务保险风险最低资本合计/认可负债"] = df_all["寿险业务保险风险最低资本合计"] / df_all["认可负债"]
    if "非寿险业务保险风险最低资本合计" in df_all.columns and "认可负债" in df_all.columns:
        df_all["非寿险业务保险风险最低资本合计/认可负债"] = df_all["非寿险业务保险风险最低资本合计"] / df_all["认可负债"]

# ===========================================================
# 侧边栏
# ===========================================================
with st.sidebar:
    st.markdown("## 寿研数智 偿付能力季报分享平台")
    st.caption("偿付能力信息可视化系统  v6.0 · PDF参照版")
    st.divider()
    
    # 数据上传区域
    st.markdown("#### 📁 数据包上传")
    uploaded_file = st.file_uploader(
        "上传 CROSS汇总表.xlsx",
        type=["xlsx"],
        key="data_uploader",
        help="请上传包含各季度数据的Excel文件"
    )
    
    if uploaded_file is not None:
        # 检查是否是新上传的文件
        if st.session_state.uploaded_file_name != uploaded_file.name:
            # 读取文件内容
            file_bytes = uploaded_file.read()
            st.session_state.uploaded_file_bytes = file_bytes
            st.session_state.uploaded_file_name = uploaded_file.name
            
            # 获取所有季度（Sheet名称）
            available_quarters = get_available_quarters(file_bytes)
            st.session_state.available_quarters = available_quarters
            
            # 设置默认季度为第一个Sheet
            if len(available_quarters) > 0:
                st.session_state.current_quarter = available_quarters[0]
            
            st.success(f"✅ 已上传：{uploaded_file.name}")
            st.rerun()
    
    # ---- 重大融资信息文件上传（可选）----
    if st.session_state.uploaded_file_name is not None:
        st.divider()
        st.markdown("#### 📁 重大融资信息（可选）")
        financing_file = st.file_uploader(
            "上传 重大融资信息.xlsx",
            type=["xlsx"],
            key="financing_uploader",
            help="用于展示04页面的重大融资信息统计，非必须"
        )
        
        if financing_file is not None:
            if st.session_state.financing_file_name != financing_file.name:
                st.session_state.financing_file_bytes = financing_file.read()
                st.session_state.financing_file_name = financing_file.name
                st.success(f"✅ 已上传：{financing_file.name}")
                st.rerun()
    
    # 显示已上传文件信息
    if st.session_state.uploaded_file_name is not None:
        st.info(f"📄 当前文件：{st.session_state.uploaded_file_name}")
        if len(st.session_state.available_quarters) > 0:
            st.caption(f"可用季度：{', '.join(st.session_state.available_quarters)}")
        # 显示融资信息文件状态
        if st.session_state.financing_file_name is not None:
            st.success(f"📄 融资信息：{st.session_state.financing_file_name}")
        else:
            st.caption("💡 未上传重大融资信息 → 04页面将跳过")
        st.divider()
    
    # 用户信息
    st.markdown(f"**当前用户：** {st.session_state.username}")
    role_label = "🔑 管理员" if st.session_state.user_role == "admin" else "👁 普通用户"
    st.caption(role_label)
    if st.button("退出登录", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.username  = ""
        st.session_state.user_role  = None
        st.rerun()
    st.divider()

    # 数据期间（从上传文件的Sheet列表中选择）
    if st.session_state.uploaded_file_name is not None and len(st.session_state.available_quarters) > 0:
        q_label = st.selectbox(
            "数据期间",
            st.session_state.available_quarters,
            index=st.session_state.available_quarters.index(st.session_state.get("current_quarter", st.session_state.available_quarters[0])) if st.session_state.get("current_quarter") in st.session_state.available_quarters else 0,
            key="quarter_selector"
        )
        # 如果季度发生变化，更新并重新加载数据
        if st.session_state.current_quarter != q_label:
            st.session_state.current_quarter = q_label
            st.rerun()
    else:
        q_label = st.text_input("数据期间", value=DEFAULT_Q, disabled=True)
        st.caption("⚠️ 请先上传数据包")
    st.divider()

    # 公司分类筛选
    st.markdown("#### 🏢 公司分类筛选")
    if df_all is not None:
        sel_cats = st.multiselect(
            "选择公司类型（留空=全部）",
            list(st.session_state.get("comp_cats", {}).keys()),
            default=[],
            key="cat_filter"
        )
        if sel_cats:
            filter_list = []
            for cat in sel_cats:
                filter_list += st.session_state.get("comp_cats", {}).get(cat, [])
            filter_list = list(dict.fromkeys(filter_list))
            df = df_all[df_all["公司"].isin(filter_list)].reset_index(drop=True)
            st.caption(f"筛选后：{len(df)} 家")
        else:
            df = df_all.copy()
            st.caption(f"全部：{len(df)} 家")
    else:
        st.caption("⚠️ 请先上传数据包后显示公司列表")
        df = None

    st.divider()

    # 目标公司
    st.markdown("#### 🎯 目标公司")
    if df is not None:
        target_co = st.selectbox(
            "追踪目标公司",
            ["（不选择）"] + df["公司"].tolist(),
            index=0
        )
        target_co = None if target_co == "（不选择）" else target_co
    else:
        target_co = None

    st.divider()

    # 导航
    st.markdown("#### 📄 页面导航")

    # 页面列表（用于 radio 和导航计算）
    page_list = [
        "01-行业整体偿付能力概览",
        "02-实际资本数据对比分析",
        "03-最低资本数据对比分析",
        "04-重大融资信息统计",
        "05-原始数据",
        "一键显示全部（打印/导出）",
    ]

    # 检查是否有外部导航请求（例如从打印按钮跳转）
    nav_target = st.session_state.get("nav_target")
    if nav_target in page_list:
        page_idx = page_list.index(nav_target)
        del st.session_state["nav_target"]
    else:
        page_idx = st.session_state.get("page_selector_idx", 0)

    page = st.radio(
        "选择页面",
        page_list,
        index=page_idx,
        key="page_selector",
        label_visibility="collapsed"
    )
    # 保存当前选择的索引，供外部导航使用
    st.session_state["page_selector_idx"] = page_list.index(page)

    st.divider()

    # 打印/导出PDF（侧边栏按钮）
    st.markdown("#### 🖨️ 打印/导出PDF")
    st.caption("点击后自动跳转到「一键显示全部」页面并打开打印对话框。")
    if st.button("16:9 横版打印", use_container_width=True, key="sidebar_print_169"):
        st.session_state.print_mode = "16-9"
        st.session_state.nav_target = "一键显示全部（打印/导出）"
        st.rerun()

    st.divider()
    st.caption(f"v5.0 · {datetime.now().strftime('%Y-%m-%d')}")
    st.caption("人身险公司偿付能力信息分享平台")

# ===========================================================
# 顶部标题
# ===========================================================
if df is not None and df_all is not None:
    gt_str = f" · 目标：{target_co}" if target_co else ""
    st.markdown(f"""
<div class="main-header">
    <h1>人身险公司偿付能力信息分享平台</h1>
    <p>中国人身险公司偿付能力季度报告数据库 · {q_label} · {len(df)} 家公司{gt_str}</p>
</div>
""", unsafe_allow_html=True)

# ===========================================================
# 页面路由
# ===========================================================

# 检查数据是否已上传
if df_all is None:
    st.warning("⚠️ 请先在左侧边栏上传数据包（CROSS汇总表.xlsx），才能查看各页面内容。")
    st.stop()

# —— 01-行业整体偿付能力概览 ——
if page == "01-行业整体偿付能力概览":
    render_page_01()
elif page == "02-实际资本数据对比分析":
    render_page_02()
elif page == "03-最低资本数据对比分析":
    render_page_03()
elif page == "04-重大融资信息统计":
    render_page_04()

# —— 一键显示全部 ——
elif page == "一键显示全部（打印/导出）":
    # 01页面（第一页，不需要手动分页）
    render_page_01(standalone=False)

    # 02页面
    render_page_02(standalone=False)

    # 03页面
    render_page_03(standalone=False)

    # 04页面
    render_page_04(standalone=False)

    # —— 打印触发器（在所有内容渲染完毕后执行）——
    if st.session_state.get("print_mode"):
        mode = st.session_state.print_mode

        # 注入完整打印 CSS + 强制 expander 展开的 JavaScript
        # 通过 iframe 注入到 parent window
        print_js = f"""
        <script>
        (function() {{
            var maxWait = 60, waited = 0;
            function checkParent() {{
                var pDoc = window.parent.document;
                if (!pDoc || !pDoc.body) {{
                    waited++;
                    if (waited < maxWait) setTimeout(checkParent, 200);
                    return;
                }}

                // 1. 注入 @page 打印尺寸
                var ps = pDoc.getElementById('dyn-page-style');
                if (!ps) {{
                    ps = pDoc.createElement('style');
                    ps.id = 'dyn-page-style';
                    pDoc.head.appendChild(ps);
                }}
                ps.innerHTML = '@page{{ size: 297mm 167mm; margin: 3mm; @bottom-right {{ content: "第 " counter(page) " 页"; font-size: 9px; color: #555; }}}}';

                // 2. 注入完整打印 CSS（隐藏侧边栏、强制 expander 展开等）
                var s = pDoc.getElementById('dyn-print-style');
                if (!s) {{
                    s = pDoc.createElement('style');
                    s.id = 'dyn-print-style';
                    pDoc.head.appendChild(s);
                }}
                s.innerHTML =
                    '@media print {{' +
                    '  [data-testid="stSidebar"],' +
                    '  [data-test-id="stSidebar"],' +
                    '  section[data-testid="stSidebar"],' +
                    '  .css-1d391kg {{' +
                    '    display: none !important; width: 0 !important;' +
                    '    min-width: 0 !important; max-width: 0 !important;' +
                    '    overflow: hidden !important; position: absolute !important;' +
                    '    left: -9999px !important;' +
                    '  }}' +
                    '  header, footer, #MainMenu {{ display: none !important; }}' +
                    '  .main .block-container {{' +
                    '    max-width: 100% !important; padding: 0 !important;' +
                    '    margin: 0 !important; width: 100% !important;' +
                    '  }}' +
                    '  .main {{ padding-left: 0 !important; margin-left: 0 !important; }}' +
                    '  .section-break {{' +
                    '    page-break-after: always !important;' +
                    '    break-after: page !important;' +
                    '    height: 0 !important; margin: 0 !important;' +
                    '    border: none !important;' +
                    '  }}' +
                    '  .js-plotly-plot, table, .stTable, ' +
                    '  [data-testid="stMetric"],' +
                    '  .element-container {{' +
                    '    page-break-inside: avoid !important;' +
                    '    break-inside: avoid !important;' +
                    '  }}' +
                    '  details, details[open],' +
                    '  [data-testid="stExpander"] details {{' +
                    '    display: block !important; visibility: visible !important;' +
                    '    height: auto !important; overflow: visible !important;' +
                    '  }}' +
                    '  details summary + div {{' +
                    '    display: block !important; visibility: visible !important;' +
                    '    height: auto !important; overflow: visible !important;' +
                    '  }}' +
                    '  .stDataFrame > div:first-child {{' +
                    '    overflow: visible !important; height: auto !important;' +
                    '    max-height: none !important;' +
                    '  }}' +
                    '  /* 强制并列栏在打印时不换行 */' +
                    '  [data-test-id="stColumns"],' +
                    '  [data-testid="stColumns"] {{' +
                    '        display: flex !important;' +
                    '        flex-wrap: nowrap !important;' +
                    '  }}' +
                    '  [data-test-id="stColumns"] > [data-test-id="stColumn"],' +
                    '  [data-testid="stColumns"] > [data-testid="stColumn"] {{' +
                    '        flex: 1 0 50% !important;' +
                    '        min-width: 50% !important;' +
                    '  }}' +
                    '  /* 打印时缩小箱型图X轴标签字体 */' +
                    '  .xtick text, g[class*="xtick"] text {{' +
                    '        font-size: 8px !important;' +
                    '  }}' +
                    '}}';
                var details = pDoc.querySelectorAll('details');
                for (var i = 0; i < details.length; i++) {{
                    details[i].setAttribute('open', '');
                }}

                // 4. 延迟执行打印，确保样式生效
                setTimeout(function() {{
                    window.parent.print();
                    // 打印后清除动态样式（可选）
                    setTimeout(function() {{
                        var s2 = pDoc.getElementById('dyn-print-style');
                        if (s2) s2.innerHTML = '';
                        var ps2 = pDoc.getElementById('dyn-page-style');
                        if (ps2) ps2.innerHTML = '';
                    }}, 1000);
                }}, 800);
            }}
            checkParent();
        }})();
        </script>
        """
        components.html(print_js, height=10, width=10, scrolling=False)

        # 打印触发后清除状态（下次 rerun 时不再重复打印）
        del st.session_state.print_mode

# —— 05-原始数据 ——
elif page == "05-原始数据":
    st.markdown(f"### 📋 05 · 原始数据浏览 · {q_label}")

    # 默认显示所有指标列
    show_cols = [c for c in df.columns if c not in ("分类",)]

    search = st.text_input("搜索公司名称", "")
    df_s = df[df["公司"].str.contains(search, na=False)] if search else df.copy()

    # 格式化显示
    df_disp = df_s[show_cols].copy() if all(c in df_s.columns for c in show_cols) else df_s.copy()
    for col in [RATIO_COMP, RATIO_CORE]:
        if col in df_disp.columns:
            df_disp[col] = df_disp[col].apply(lambda v: pct(v) if pd.notna(v) else "—")
    for col in df_disp.columns:
        if col not in ["公司",RATIO_COMP,RATIO_CORE] and pd.api.types.is_numeric_dtype(df_disp[col]):
            df_disp[col] = df_disp[col].apply(lambda v: f"{v:,.0f}" if pd.notna(v) and abs(float(v))>=10 else (f"{v:.4f}" if pd.notna(v) else "—"))

    st.markdown(f"共 **{len(df_s)}** 家公司 · **{len(df_disp.columns)-1}** 个指标")
    st.dataframe(df_disp, use_container_width=True, hide_index=True, height=600)

    # 下载
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_s[show_cols].to_excel(w, index=False, sheet_name="CROSS数据")
    buf.seek(0)
    st.download_button(
        "⬇ 下载当前数据（Excel）",
        data=buf,
        file_name=f"CROSS_{q_label}_导出.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
