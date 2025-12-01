import streamlit as st

def load_css():
    st.markdown("""
        <style>
        /* 全局字体优化 */
        .stApp {
            font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
        }

        /* 统计容器 */
        .stat-container { 
            background-color: #ffffff; 
            padding: 20px; 
            border-radius: 12px; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.08); 
            border: 1px solid #f1f5f9;
        }
        .stat-header { 
            color: #1e293b; 
            font-size: 1.2em; 
            font-weight: 700; 
            margin-bottom: 20px; 
            border-bottom: 2px solid #e2e8f0; 
            padding-bottom: 10px; 
            display: flex;
            align-items: center;
        }
        .stat-header:before {
            content: "";
            margin-right: 8px;
            font-size: 1.1em;
        }
        
        /* 汇总卡片网格 */
        .summary-grid { 
            display: grid; 
            grid-template-columns: repeat(2, 1fr); 
            gap: 12px; 
            margin-bottom: 24px; 
        }
        .summary-card { 
            background: linear-gradient(145deg, #ffffff, #f8fafc); 
            padding: 12px; 
            border-radius: 10px; 
            text-align: center; 
            border: 1px solid #e2e8f0; 
            transition: all 0.3s ease; 
            position: relative;
            overflow: hidden;
        }
        .summary-card:hover { 
            transform: translateY(-3px); 
            box-shadow: 0 8px 16px rgba(0,0,0,0.06); 
            border-color: #cbd5e1; 
        }
        .summary-card:after {
            content: "";
            position: absolute;
            top: 0; left: 0; width: 100%; height: 4px;
            background: #3b82f6;
            opacity: 0.8;
        }
        .summary-val { 
            font-size: 1.6em; 
            font-weight: 800; 
            color: #0f172a; 
            line-height: 1.2; 
            margin-top: 4px;
        }
        .summary-label { 
            font-size: 0.85em; 
            color: #64748b; 
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        /* 层级列表样式 */
        .module-l1 { 
            background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%); 
            color: white; 
            padding: 12px 16px; 
            border-radius: 8px; 
            margin: 20px 0 8px 0; 
            font-weight: 700; 
            font-size: 1em;
            box-shadow: 0 4px 6px rgba(37, 99, 235, 0.25);
            display: flex; 
            align-items: center;
            letter-spacing: 0.3px;
        }
        .module-l1-icon { margin-right: 10px; font-size: 1.1em; }
        
        .module-l2 { 
            background-color: #f1f5f9; 
            color: #334155; 
            padding: 10px 14px; 
            border-radius: 6px; 
            margin: 6px 0 6px 16px; 
            border-left: 4px solid #64748b; 
            font-weight: 600; 
            font-size: 0.95em;
            display: flex;
            align-items: center;
        }
        .module-l2:before {
            content: "";
            margin-right: 8px;
            font-size: 0.9em;
            opacity: 0.7;
        }

        .module-l3 {
            background-color: #ffffff;
            color: #475569;
            padding: 8px 12px;
            border-radius: 6px;
            margin: 4px 0 4px 32px;
            border: 1px solid #e2e8f0;
            font-size: 0.85em;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .module-l3:hover {
            border-color: #cbd5e1;
            background-color: #f8fafc;
        }
        .module-l3-count {
            background-color: #eff6ff;
            color: #3b82f6;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 0.85em;
            font-weight: 600;
        }
        
        /* 三级模块 Expander 优化 */
        .streamlit-expanderHeader {
            background-color: #ffffff !important;
            border: 1px solid #e2e8f0 !important;
            border-radius: 6px !important;
            font-size: 0.8em !important;
            color: #475569 !important;
            font-weight: 500 !important;
            margin-left: 24px !important;
            margin-bottom: 2px !important;
            transition: background-color 0.2s;
        }
        
        /* 强制覆盖 Streamlit 内部 p 标签样式 */
        .streamlit-expanderHeader p {
            font-size: 0.85em !important;
            font-weight: 500 !important;
        }
        .streamlit-expanderHeader:hover {
            background-color: #f8fafc !important;
            color: #2563eb !important;
            border-color: #cbd5e1 !important;
        }
        .streamlit-expanderContent {
            margin-left: 24px !important;
            border-left: 2px solid #e2e8f0 !important;
            padding-left: 12px !important;
            padding-bottom: 8px !important;
        }

        /* 功能过程列表项 */
        .process-item {
            background-color: transparent;
            padding: 4px 8px;
            margin: 2px 0;
            border-radius: 4px;
            font-size: 0.85em;
            color: #64748b;
            display: flex; 
            align-items: center;
            transition: color 0.2s;
        }
        .process-item:hover {
            color: #2563eb;
            background-color: #f8fafc;
        }
        .process-item:before {
            content: "•";
            color: #cbd5e1;
            margin-right: 8px;
            font-size: 1.4em;
            line-height: 0;
            position: relative;
            top: 1px;
        }
        
        /* 按钮美化 */
        div.stButton > button {
            border-radius: 8px;
            font-weight: 600;
            transition: all 0.2s;
        }
        div.stButton > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        </style>
    """, unsafe_allow_html=True)
