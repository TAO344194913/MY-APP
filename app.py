import streamlit as st
import re

def main():
    st.title("📋 多专业名单汇总系统")
    st.markdown("""
    **使用说明：**
    1. 在下方输入框按格式输入专业和名单
    2. 格式示例：`计算机专业：张三、李四、王五`
    3. 可以一次性输入多行，用换行分隔不同专业
    4. 点击"提交"按钮处理所有输入
    """)

    # 初始化session state保存数据
    if 'major_dict' not in st.session_state:
        st.session_state.major_dict = {}
    if 'all_inputs' not in st.session_state:
        st.session_state.all_inputs = ""

    # 多行文本输入框
    inputs = st.text_area("请输入专业和名单（可多行输入）", 
                        height=150,
                        placeholder="例如：\n计算机专业：张三、李四\n数学专业：王五、赵六\n...")

    # 提交按钮
    if st.button("🚀 提交处理"):
        if inputs:
            process_inputs(inputs)
            st.session_state.all_inputs = inputs  # 保存当前输入
        else:
            st.warning("请输入内容后再提交")

    # 显示处理结果
    show_results()

def process_inputs(input_text):
    """处理多行输入"""
    pattern = r"([\u4e00-\u9fa5]+专业)[：:]\s*([\u4e00-\u9fa5]+(?:、[\u4e00-\u9fa5]+)*)"
    lines = input_text.split('\n')  # 按行分割
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        matches = re.findall(pattern, line)
        if matches:
            major, names = matches[0]
            name_list = names.split('、')
            
            # 更新专业字典
            if major in st.session_state.major_dict:
                st.session_state.major_dict[major].extend(name_list)
            else:
                st.session_state.major_dict[major] = name_list

def show_results():
    """显示所有汇总结果"""
    st.subheader("📊 专业名单汇总")
    
    if st.session_state.major_dict:
        # 按专业分类显示
        for major, students in st.session_state.major_dict.items():
            with st.expander(f"{major} (共{len(students)}人)"):
                cols = st.columns(4)
                for i, name in enumerate(students):
                    cols[i%4].write(f"• {name}")
        
        # 显示原始输入
        st.subheader("📝 原始输入记录")
        st.text(st.session_state.all_inputs)
    else:
        st.info("尚未输入任何专业信息")

    # 添加清空按钮
    if st.button("🧹 清空所有数据"):
        st.session_state.major_dict = {}
        st.session_state.all_inputs = ""
        st.experimental_rerun()

if __name__ == "__main__":
    main()
