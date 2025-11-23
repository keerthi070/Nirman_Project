import streamlit as st
from scorer import score_transcript, read_rubric_from_excel
import json
import os

# ---------------------------------------------------
# Streamlit Page Setup
# ---------------------------------------------------
st.set_page_config(page_title="Nirmaan AI – Transcript Scorer", layout="wide")
st.title("🎤 Student Introduction Evaluator – Nirmaan AI Internship")

# ---------------------------------------------------
# Load Rubric Once
# ---------------------------------------------------
RUBRIC_PATH = "Case study for interns.xlsx"

try:
    rubric = read_rubric_from_excel(RUBRIC_PATH)
    st.success("Rubric loaded successfully.")
except:
    st.warning("Could not load rubric Excel. Using fallback rubric.")
    rubric = read_rubric_from_excel(None)

# ---------------------------------------------------
# Input Section
# ---------------------------------------------------
st.subheader("📥 Enter Transcript Text")

text_input = st.text_area(
    "Paste the transcript here:",
    height=250,
)

uploaded_file = st.file_uploader("Or upload transcript (.txt)", type=["txt"])

if uploaded_file:
    text_input = uploaded_file.read().decode("utf-8")


# ---------------------------------------------------
# Button with session_state to avoid double execution
# ---------------------------------------------------
if st.button("🔍 Evaluate Transcript"):
    st.session_state["run"] = True
    st.session_state["input_text"] = text_input

if st.session_state.get("run", False):

    transcript = st.session_state.get("input_text", "").strip()

    if not transcript:
        st.error("❌ Please enter or upload a transcript.")
    else:
        results = score_transcript(transcript, rubric)

        st.success(f"✅ Overall Score: {results['overall_score']} / 100")
        st.markdown("---")

        # ---------------------------------------------------
        # Display Detailed Scores (No Duplicate Output)
        # ---------------------------------------------------
        st.subheader("📊 Detailed Criterion Scores")

        shown = set()

        for criterion, info in results["criterion_scores"].items():

            if criterion in shown:
                continue
            shown.add(criterion)

            st.markdown(f"""
            ### 🟦 {criterion}
            **Score:** {info['score']} / 100  
            **Weight:** {info['weight']}  
            **Rule Score:** {info['rule_score']}  
            **Semantic Score:** {info['semantic_score']}  
            **Keywords Found:** {info['keywords_found']}  
            **Word Count:** {info['num_words']}  
            """)

        st.markdown("---")

        # Download JSON
        st.download_button(
            label="⬇ Download JSON Report",
            file_name="score_output.json",
            mime="application/json",
            data=json.dumps(results, indent=4)
        )
