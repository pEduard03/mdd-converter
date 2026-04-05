import streamlit as st
import docx
import google.generativeai as genai

# 1. Configurarea paginii web
st.set_page_config(page_title="Word to MDD Converter", page_icon="📝", layout="wide")

st.title("📝 Word to MDD/Dimensions Script Converter")
st.markdown("Încarcă un fișier Word cu chestionarul tău.")

# 2. Input pentru API Key
api_key = st.text_input("Introdu cheia", type="password")

# 3. Funcția de citire Word
def extract_text_from_docx(file):
    doc = docx.Document(file)
    content = []
    
    for para in doc.paragraphs:
        if para.text.strip():
            content.append(para.text.strip())
            
    content.append("\n--- TABELE ---")
    for table in doc.tables:
        for row in table.rows:
            row_data = [cell.text.replace('\n', ' ').strip() for cell in row.cells if cell.text.strip()]
            seen = set()
            unique_data = [x for x in row_data if not (x in seen or seen.add(x))]
            if unique_data:
                content.append(" | ".join(unique_data))
                
    text_final = "\n".join(content)
    return text_final

# 4. Încărcarea fișierului Word
uploaded_file = st.file_uploader("Încarcă fișierul .docx", type=["docx"])

# 5. Logica principală de procesare
if uploaded_file is not None and api_key:
    if st.button("Generează Codul MDD"):
        with st.spinner("Citesc documentul și generez scriptul..."):
            try:
                document_text = extract_text_from_docx(uploaded_file)
                
                if len(document_text) < 10:
                    st.error("Documentul pare să fie gol sau formatul nu a putut fi citit corect.")
                    st.stop()
                
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-2.5-flash') 
                
                # NOUL PROMPT - CU REGULILE PENTRU SPAN, DIV, FIX ȘI RAN ACTUALIZATE
                prompt = f"""
                You are an Expert Survey Programmer writing MDD/Dimensions script.
                Convert the raw survey text into COMPLETE MDD code. Do NOT summarize.

                CRITICAL SYNTAX RULES (DO NOT BREAK):
                1. IGNORE HEADERS: Ignore text like "LINEARES TV", "STREAMING". Only process actual questions (e.g., 1., 2., Q1).
                2. QUOTES ESCAPING: NEVER use backslashes (`\\"`) for HTML. You MUST use double-double quotes. 
                   CORRECT: `<div class=""qtext"">`
                   WRONG: `<div class=\\"qtext\\">`
                3. SELECTIONS: Use `categorical [0..1]` for Single Choice and `categorical [0..]` (or `[0..3]`) for Multiple Choice. NEVER append `max(3)` at the end of the block.
                4. OTHER/OPEN TEXT: Use the exact syntax `other(other_id "" text[0..]) fix` for "Sonstiges" options.
                5. EXCLUSIVE CODES & RANDOMIZATION: Add `fix exclusive` to "weiß nicht", "keine", etc. NEVER use `ran except` syntax. Just use `ran;` at the end of the block. The `fix` keyword already handles the anchoring.
                6. GRID SYNTAX: The Question Name and Text MUST come BEFORE the `loop` keyword. DO NOT output `SkipSecurity = "Yes"`.
                7. GRID FIELDS: ALWAYS use `_scale ""` inside the `fields -` block. NEVER invent names like `_providers ""`.
                8. PIPING / VARIABLES: Replace placeholders like "[show kids name]" or "[show kids name in blue letters]" with the exact MDD piping syntax: `{{#kids_name.response.value}}`. If a color is specified in the placeholder (like blue), wrap the piping in HTML: `<span style='color:blue'><strong>{{#kids_name.response.value}}</strong></span>`.
                9. NO ROUTING LOGIC: DO NOT output any routing logic, filter conditions, or 'VisibleIf' statements (e.g., VisibleIf = ...). Keep it strictly structural.
                10. BANKED GRIDS: ALL grids/loops MUST include `GfKGridType = "banked"` inside the square brackets property block `[ ... ]`.
                11. HTML FORMATTING FOR TEXTS: 
                    - The main question text MUST be wrapped in `<strong>...</strong>`.
                    - The `<div class=""qtext"">` tag MUST close immediately after `</strong>`. Example: `<div class=""qtext""><strong>Question?</strong></div>`
                    - Any instruction text (e.g., "Bitte geben Sie alles Zutreffende an.", "Mehrfachantworten möglich") MUST be placed OUTSIDE the div and wrapped in `<span class=""instruction"">...</span>`.

                --- REQUIRED MULTI-CHOICE TEMPLATE ---
                '=====Q1 Base: all respondents; randomize items except code 99, [M]; [SC]
                Q1 "<div class=""qtext""><strong>Welche der genannten Geräte sind in Ihrem Haushalt vorhanden?</strong></div><span class=""instruction"">Bitte geben Sie alles Zutreffende an.</span>"
                categorical [0..]
                {{
                    _1 "Smartphone",
                    _99 "keines der genannten" fix exclusive
                }} ran;

                --- REQUIRED GRID TEMPLATE (COPY THIS EXACT STRUCTURE) ---
                '=====Q2 Base: all respondents, randomize items [S per row], [HC per row]
                Q2 "<div class=""qtext""><strong>Im Folgenden sehen Sie verschiedene Medien. Bitte geben Sie an, wie oft Sie diese privat nutzen.</strong></div>"
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

                --- REQUIRED SUBLIST TEMPLATE (INSIDE LOOPS) ---
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

                ONLY output the raw MDD code based on these exact templates. No markdown formatting outside the code block.

                Raw Survey Document to convert:
                \n\n{document_text}
                """
                
                response = model.generate_content(prompt)
                
                try:
                    final_code = response.text
                    st.success("Conversie finalizată!")
                    st.code(final_code, language="mdd")
                except Exception as eval_error:
                    st.error("Am întâmpinat o eroare la afișare.")
                    st.write(response.candidates[0])
                
            except Exception as e:
                st.error(f"A apărut o eroare generală: {e}")
                
elif not api_key and uploaded_file:
    st.warning("Te rog să introduci cheia API Gemini pentru a continua.")