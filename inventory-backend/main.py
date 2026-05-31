from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Dict, Any, List, Optional
import json
import re
from datetime import date
from ai_client import call_ai
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ========== Models ==========
class ChatMessage(BaseModel):
    role: str
    content: str

class ProcessRequest(BaseModel):
    user_input: str
    sheet_schema: Dict[str, Any]
    chat_history: List[ChatMessage] = []

class ProcessResponse(BaseModel):
    actions: List[Dict[str, Any]]
    reply: str
    warnings: List[str]

# ========== Helpers ==========
def extract_json_from_response(text: str) -> dict:
    cleaned = re.sub(r'```json\s*', '', text)
    cleaned = re.sub(r'```\s*', '', cleaned)
    start = cleaned.find('{')
    end = cleaned.rfind('}')
    if start == -1 or end == -1:
        raise ValueError("No JSON object found in AI response")
    json_str = cleaned[start:end+1]
    return json.loads(json_str)

# ========== Main Endpoint ==========
@app.post("/process", response_model=ProcessResponse)
async def process_natural_language(request: ProcessRequest):
    try:
        today = date.today().strftime("%d/%m/%Y")
        history_text = "\n".join(
            [f"{msg.role.upper()}: {msg.content}" for msg in request.chat_history[-20:]]
        )

        schema_text = ""
        for sheet, headers in request.sheet_schema.items():
            clean_headers = [str(h) for h in headers if h]
            schema_text += f'  Sheet "{sheet}": {", ".join(clean_headers) if clean_headers else "no headers yet"}\n'

        system_prompt = f"""
You are a universal Excel AI assistant. You work with ANY spreadsheet for ANY industry.
Industries you handle: inventory, banking, HR, projects, hospitals, schools, retail, finance, and more.

WORKBOOK STRUCTURE:
{schema_text}

CONVERSATION HISTORY:
{history_text}

TODAY: {today}

YOUR CAPABILITIES AND ACTION TYPES:

1. ADD ROW → type: "add_row"
   Use when user says: add, insert, new entry, record, log
   Fields: "sheet" (exact name), "data" (object with column:value pairs)
   Example: {{"type":"add_row","sheet":"Employees","data":{{"Name":"John","Salary":50000,"Dept":"Sales"}}}}

2. EDIT ROW → type: "edit_row"
   Use when user says: update, change, modify, fix, correct
   Fields: "sheet", "search" (column:value to find), "new_values" (columns to change)
   Example: {{"type":"edit_row","sheet":"Employees","search":{{"Name":"John"}},"new_values":{{"Salary":55000}}}}

3. DELETE ROW → type: "delete_row"
   Use when user says: delete, remove, erase
   Fields: "sheet", "search" (column:value to find row)
   Example: {{"type":"delete_row","sheet":"Employees","search":{{"Name":"John"}}}}

4. INSERT FORMULA → type: "formula"
   Use when user says: formula, calculate, sum, average, count, IF, VLOOKUP
   Fields: "sheet", "cell" (like "E2"), "formula" (starts with =)
   Example: {{"type":"formula","sheet":"Sales","cell":"D2","formula":"=SUM(B2:C2)"}}

5. FORMAT CELLS → type: "format"
   Use when user says: bold, color, highlight, format, make red, fill
   Fields: "sheet", "range" (like "A1:D1"), "formatting" (object)
   Formatting options: bold (true/false), color (hex font color), fillColor (hex bg), fontSize (number)
   Example: {{"type":"format","sheet":"Sales","range":"A1:E1","formatting":{{"bold":true,"fillColor":"#1a1a2e"}}}}

6. CLEAN COLUMN → type: "clean"
   Use when user says: clean, trim, fix spaces, uppercase, lowercase, standardize
   Fields: "sheet", "column" (exact header name), "clean_action" (trim/uppercase/lowercase)
   Example: {{"type":"clean","sheet":"Employees","column":"Name","clean_action":"trim"}}

7. CREATE SHEET → type: "create_sheet"
   Use when user says: create sheet, new sheet, add tab
   Fields: "new_sheet" (name), "headers" (array of column names)
   Example: {{"type":"create_sheet","new_sheet":"Q2 Report","headers":["Date","Item","Amount"]}}

8. SORT DATA → type: "sort"
   Use when user says: sort, order, arrange by
   Fields: "sheet", "sort_by" (column name), "order" (asc/desc)
   Example: {{"type":"sort","sheet":"Sales","sort_by":"Date","order":"desc"}}

9. ANSWER QUESTION → type: "query"
   Use when user asks a question with no data change needed
   No extra fields. Just reply conversationally.

CRITICAL RULES:
- Match sheet names EXACTLY as shown in WORKBOOK STRUCTURE (case sensitive)
- Match column names EXACTLY as shown
- If user misspells a name, intelligently correct it to the closest match
- If sheet/column not clear, ask in reply and use type "query"
- Handle multi-step: one message can produce multiple actions in the array
- Always confirm what you did in reply field
- RETURN ONLY VALID JSON — no markdown, no explanation outside the JSON

RETURN FORMAT (always this exact structure):
{{
  "actions": [ {{...}}, {{...}} ],
  "reply": "Natural conversational confirmation of what was done or answer to question",
  "warnings": ["optional warnings like low stock, missing data, etc"]
}}

User said: "{request.user_input}"
JSON response:
"""

        ai_response = call_ai(system_prompt)
        print(f"\n--- AI RAW RESPONSE ---\n{ai_response}\n---")

        parsed = extract_json_from_response(ai_response)

        actions = parsed.get("actions", [])
        reply = parsed.get("reply", "Done.")
        warnings = parsed.get("warnings", [])

        # Validate each action has required fields
        validated_actions = []
        for action in actions:
            atype = action.get("type", "")
            if atype == "add_row" and ("sheet" not in action or "data" not in action):
                warnings.append(f"Skipped add_row — missing sheet or data")
                continue
            if atype in ["edit_row", "delete_row"] and "sheet" not in action:
                warnings.append(f"Skipped {atype} — missing sheet")
                continue
            if atype == "formula" and not all(k in action for k in ["sheet", "cell", "formula"]):
                warnings.append("Skipped formula — missing sheet, cell, or formula")
                continue
            if atype == "format" and not all(k in action for k in ["sheet", "range", "formatting"]):
                warnings.append("Skipped format — missing sheet, range, or formatting")
                continue
            if atype == "clean" and not all(k in action for k in ["sheet", "column"]):
                warnings.append("Skipped clean — missing sheet or column")
                continue
            if atype == "create_sheet" and "new_sheet" not in action:
                warnings.append("Skipped create_sheet — missing sheet name")
                continue
            if atype == "sort" and not all(k in action for k in ["sheet", "sort_by"]):
                warnings.append("Skipped sort — missing sheet or sort_by")
                continue
            validated_actions.append(action)

        return ProcessResponse(actions=validated_actions, reply=reply, warnings=warnings)

    except json.JSONDecodeError as e:
        return ProcessResponse(
            actions=[],
            reply="I had trouble formatting my response. Please rephrase your request.",
            warnings=[str(e)]
        )
    except Exception as e:
        return ProcessResponse(
            actions=[],
            reply=f"Something went wrong: {str(e)}. Please try again.",
            warnings=[]
        )

@app.get("/health")
async def health():
    return {"status": "alive", "date": date.today().strftime("%d/%m/%Y")}

# ========== Static Files — MUST BE LAST ==========
static_dir = os.path.join(os.path.dirname(__file__), "..", "excel-addin")
if os.path.exists(static_dir):
    app.mount("/", StaticFiles(directory=static_dir, html=True), name="static")
    print(f"[OK] Serving frontend from: {static_dir}")
else:
    print(f"[WARN] Frontend folder not found at: {static_dir}")