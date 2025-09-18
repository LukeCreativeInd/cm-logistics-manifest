import streamlit as st

st.set_page_config(page_title="CM Logistics Manifest Generator", layout="centered")

# Display logo (optional)
st.image("CM_Logistics_Top_Logo.png", width="stretch")

st.title("CM Logistics Manifest Generator")
st.markdown("### Which customer group do you want to create a manifest for?")

# Initialize session state
if "selected_group" not in st.session_state:
    st.session_state.selected_group = None

# Horizontal buttons
col1, col2, col3 = st.columns(3)

with col1:
    if st.button("Clean Eats Australia"):
        st.session_state.selected_group = "Clean Eats Australia"

with col2:
    if st.button("Made Active"):
        st.session_state.selected_group = "Made Active"

with col3:
    if st.button("Elite Meals"):
        st.session_state.selected_group = "Elite Meals"

# Render group UI at full width
selected = st.session_state.selected_group

if selected == "Clean Eats Australia":
    import clean_eats
    clean_eats.run()

elif selected == "Made Active":
    import made_active
    made_active.run()

elif selected == "Elite Meals":
    import elite_meals
    elite_meals.run()
