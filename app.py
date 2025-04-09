import streamlit as st

# 标题
st.title("我的第一个网页程序")

# 输入框
name = st.text_input("请输入你的名字")

# 按钮
if st.button("打招呼"):
    st.write(f"你好，{name}！")