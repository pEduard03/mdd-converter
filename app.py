import re
import streamlit as st
import docx
import google.generativeai as genai


# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="Word to MDD Converter",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Word to MDD/Dimensions Script Converter")
st.markdown("Încarcă un fișier Word cu chestionarul tău.")


# =========================
# API KEY INPUT
# =========================
api_key = st.text_input("Introdu cheia ", type="password")


# =========================
# HELPERS
# =========================
def extract_text_from_docx(file) -> str:
    doc = docx.Document(file)
    content = []

    # Paragraphs
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            content.append(text)

    # Tables
    table_lines = []
    for table in doc.tables:
        for row in table.rows:
            row_data = []

            for cell in row.cells:
                cell_text = " ".join(
                    line.strip() for line in cell.text.splitlines() if line.strip()
                ).strip()

                if cell_text:
                    row_data.append(cell_text)

            unique_data = []
            for item in row_data:
                if item not in unique_data:
                    unique_data.append(item)

            if unique_data:
                table_lines.append(" | ".join(unique_data))

    if table_lines:
        content.append("\n--- TABELE ---")
        content.extend(table_lines)

    return "\n".join(content).strip()


def clean_model_output(text: str) -> str:
    if not text:
        return ""

    cleaned = text.strip()
    cleaned = re.sub(r"^```[a-zA-Z0-9_-]*\s*", "", cleaned)
    cleaned = re.sub(r"\s*```$", "", cleaned)

    return cleaned.strip()


def get_response_text(response) -> str:
    try:
        if hasattr(response, "text") and response.text:
            return response.text
    except Exception:
        pass

    try:
        parts = []
        for candidate in response.candidates:
            for part in candidate.content.parts:
                if hasattr(part, "text") and part.text:
                    parts.append(part.text)
        return "\n".join(parts).strip()
    except Exception:
        return ""


def build_prompt(document_text: str) -> str:
    return f"""
You are an Expert Survey Programmer writing MDD/Dimensions script.
Convert the raw survey text into COMPLETE MDD code.
Do NOT summarize.
Do NOT explain.
Output ONLY raw MDD code.

CRITICAL SYNTAX RULES (DO NOT BREAK):
1. IGNORE SECTION HEADERS:
   Ignore non-question headers and noise such as:
   - "LINEARES TV"
   - "STREAMING"
   - page numbers
   - tester notes
   - wave notes
   - "not relevant for scripting"
   - background notes
   Only process actual survey questions and their answer structures.

2. DO NOT ADD HIDDEN QUESTIONS:
   Do NOT create hidden questions such as StatusCP.
   Do NOT invent helper variables, counters, or metadata questions unless they are explicit survey questions in the source.

3. PRESERVE ORIGINAL CODES EXACTLY:
   NEVER renumber, normalize, compress, or reorder answer codes.
   Keep original numeric codes exactly as they appear in the source, including non-sequential codes like 14, 21, 24, 99.

4. QUOTES ESCAPING:
   NEVER use backslashes (\\" ) for HTML.
   You MUST use double-double quotes.
   CORRECT: <div class=""qtext"">
   WRONG: <div class=\\"qtext\\">

5. SELECTIONS:
   - Use categorical [0..1] for Single Choice.
   - Use categorical [0..] or categorical [0..3] for Multiple Choice.
   - NEVER append max(3) or similar text after the categorical block.

6. OTHER / OPEN TEXT MUST STAY ON THE SAME LINE:
   For "Sonstiges" / "Other" style options, use the exact syntax:
   other(other_id "" text[0..]) fix
   and keep it on the SAME LINE as the statement.
   CORRECT:
   _13 "Sonstiges, und zwar:" other(other_13 "" text[0..]) fix

7. EXCLUSIVE CODES & RANDOMIZATION:
   Add fix exclusive to options such as:
   - "weiß nicht"
   - "keines der genannten"
   - "keinen davon"
   NEVER use ran except.
   Just use ran; at the end of the block.

8. GRID SYNTAX:
   The Question Name and Text MUST come BEFORE the loop keyword.
   DO NOT output SkipSecurity = "Yes".

9. GRID FIELDS:
   - For normal categorical grids, ALWAYS use _scale "" inside the fields - block.
   - NEVER invent field names like _providers "".
   - For simple open-end list grids, also use _scale "" but with text[0..]; instead of categorical.

10. PIPING / VARIABLES:
   Replace placeholders like:
   - [show kids name]
   - [show kids name in blue letters]
   with exact MDD piping syntax:
   {{{{#kids_name.response.value}}}}
   If color is specified, use:
   <span style='color:blue'><strong>{{{{#kids_name.response.value}}}}</strong></span>

11. NO ROUTING LOGIC:
   DO NOT output routing logic, filter conditions, masks, helper derivations, or VisibleIf.
   Keep it strictly structural MDD only.

12. BANKED GRIDS:
   ALL normal categorical grids/loops MUST include:
   [
       GfKGridType = "banked"
   ]
   EXCEPTION:
   Do NOT use banked for simple repeated OE/title-entry grids.

13. HTML FORMATTING FOR QUESTION TEXTS:
   - NEVER automatically wrap the whole question text in <strong>...</strong>.
   - Preserve bold EXACTLY where it is intended from the source text.
   - Only the words that are actually emphasized in the source may be wrapped in <strong>.
   - If no words are explicitly emphasized in the source, do NOT invent bold text.
   - Keep all question text inside:
     <div class=""qtext""> ... </div>

14. HTML FORMATTING FOR INSTRUCTIONS:
   - Any instruction text MUST be outside the qtext div.
   - The instruction span MUST start on the NEXT ROW after </div>, never on the same row.
   - Always format instruction text exactly like this:

     QX "<div class=""qtext"">Question text here</div>
     <span class=""instruction"">Instruction text here</span>"

   - NEVER output:
     </div><span class=""instruction"">...</span>

15. DO NOT CHANGE QUESTION / ANSWER MEANING:
   Do not paraphrase unless needed for correct syntax.
   Do not invent providers, genres, sublists, answer categories, texts, or codes.

16. NUMERIC / OE TYPES:
   - If the source clearly says [N], NEVER output numeric;
   - For [N] questions, ALWAYS output text[0..];
   - If the source clearly says [OE], output open text syntax;
   - Keep structure faithful to source.

17. SUBLISTS:
   If the source clearly groups loop items into sublists, preserve them as sublist blocks.
   Example:
   loop
   {{
       sublist1 ""
       {{
           _1 "RTL",
           _2 "VOX"
       }} ran,
       sublist2 ""
       {{
           _3 "SAT.1"
       }} ran
   }} ran fields - ...

18. SIMPLE REPEATED OE LIST QUESTIONS:
   If a question is an open-end question [OE] with multiple blank answer lines like:
   1 __________
   2 __________
   3 __________
   4 __________
   5 __________
   then convert it into a SIMPLE GRID / LOOP WITHOUT BANKED.

   IMPORTANT:
   - Do NOT add [ GfKGridType = "banked" ]
   - Do NOT use categorical
   - Use loop with one statement per blank line
   - Keep the real number of answer lines from the source
   - In fields -, use exactly:
     _scale ""
     text[0..];

19. DO NOT USE MARKDOWN:
   Do NOT wrap output in triple backticks.
   Do NOT add commentary before or after the code.

REFERENCE PATTERNS:

'=====Multiple choice with instruction on next row
Q1 "<div class=""qtext"">Welche der genannten Geräte sind in Ihrem Haushalt vorhanden?</div>
<span class=""instruction"">Bitte geben Sie alles Zutreffende an.</span>"
categorical [0..]
{{
    _1 "Smartphone",
    _99 "keines der genannten" fix exclusive
}} ran;

'=====Grid with banked
Q2 "<div class=""qtext"">Im Folgenden sehen Sie verschiedene Medien. Bitte geben Sie an, wie oft Sie diese privat nutzen.</div>"
[
    GfKGridType = "banked"
]
loop
{{
    _1 "Klassisches (lineares) Fernsehen",
    _2 "Online-Angebote"
}} ran fields -
(
    _scale ""
    categorical [0..1]
    {{
        _1 "täglich",
        _2 "mehrmals die Woche"
    }};
) expand grid;

'=====Example with partial bold only where intended
Q15 "<div class=""qtext"">Auf welchen Produktkategorien hat <strong>Hofer</strong> aktuell zu wenige/unattraktive Aktionen / Angebote?</div>
<span class=""instruction"">Bitte wählen Sie <strong>bis zu 5 Produktkategorien</strong> aus!</span>"
categorical [0..5]
{{
    _1 "Kategorie 1",
    _2 "Kategorie 2",
    _3 "Kategorie 3"
}} ran;

'=====Repeated OE list example - NO BANKED
Q14 "<div class=""qtext"">Nenne uns deine Lieblingsserien oder -filme oder auch Lieblingssendungen beim Fernsehen und Streaming.</div>
<span class=""instruction"">Du kannst bis zu 5 Titel angeben.</span>"
loop
{{
    _1 "1",
    _2 "2",
    _3 "3",
    _4 "4",
    _5 "5"
}} fields -
(
    _scale ""
    text[0..];
) expand grid;

Raw Survey Document to convert:

{document_text}
""".strip()

# =========================
# FILE UPLOAD
# =========================
uploaded_file = st.file_uploader("Încarcă fișierul .docx", type=["docx"])


# =========================
# MAIN LOGIC
# =========================
if uploaded_file is not None and api_key:
    if st.button("Generează Codul MDD", type="primary"):
        with st.spinner("Citesc documentul și generez scriptul..."):
            try:
                document_text = extract_text_from_docx(uploaded_file)

                if len(document_text) < 10:
                    st.error("Documentul pare să fie gol sau formatul nu a putut fi citit corect.")
                    st.stop()

                genai.configure(api_key=api_key)
                model = genai.GenerativeModel("gemini-2.5-flash")

                prompt = build_prompt(document_text)
                response = model.generate_content(prompt)

                final_code = clean_model_output(get_response_text(response))

                if not final_code:
                    st.error("Modelul nu a returnat text valid.")
                    st.write(response)
                    st.stop()

                st.success("Conversie finalizată!")
                st.code(final_code, language="text")

                st.download_button(
                    label="Descarcă scriptul MDD",
                    data=final_code,
                    file_name="survey_script.mdd",
                    mime="text/plain"
                )

            except Exception as e:
                st.error(f"A apărut o eroare generală: {e}")

elif uploaded_file is not None and not api_key:
    st.warning("Te rog să introduci cheia ")