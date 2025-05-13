import streamlit as st

st.set_page_config(page_title="CM Logistics Manifest Generator", layout="centered")

# Display logo (optional)
st.image("CM_Logistics_Top_Logo.png", use_container_width=True)

# Title & Instructions
st.title("CM Logistics Manifest Generator")
st.markdown("### Which customer group do you want to create a manifest for?")

# Horizontal buttons
col1, col2, col3 = st.columns(3)

with col1:
    if st.button("Clean Eats Australia"):
        import clean_eats
        clean_eats.run()

with col2:
    if st.button("Made Active"):
        import made_active
        made_active.run()

with col3:
    if st.button("Elite Meals"):
        import elite_meals
        elite_meals.run()
