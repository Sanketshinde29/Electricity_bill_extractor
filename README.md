# Electricity_bill_extractor


Use your own key  as Groq Key:
 

Upload image → Base64 encode → Groq AI reads → JSON extract → Excel download

User uploads a bill photo → Streamlit sends it to Groq's Llama 4 Scout vision model → AI returns structured JSON with all bill fields → openpyxl builds a formatted Excel → user downloads the .xlsx file instantly.

I developed a Streamlit-based web application that uses AI-powered image processing to extract structured data from electricity bills and convert it into a usable format.

Technologies & Tools Used
Frontend & App Framework
Streamlit — for building the interactive web interface
Backend & Logic
Python — core programming language
SQLite — lightweight local database for storing bill history
AI Integration
Groq API with LLaMA 4 Scout model
Used for extracting structured JSON data from bill images
File Handling
base64 — for encoding images before sending to AI
io — handling in-memory file streams
Excel Generation
openpyxl — for creating styled Excel reports
Includes formatting (colors, borders, alignment)


Future Optimization:
1.Bill history — save past extractions
  Store extracted data in SQLite so users can view, compare, and re-download past bills without re-uploading images.
2.Multi-language bills — Hindi, Marathi, Tamil
  Adjust the system prompt to handle regional language bills. Groq's Llama 4 model handles Devanagari and other Indian scripts natively.
3.Smart alerts — detect unusual bills
  Compare current bill's units with the past 3-month average. Automatically flag if units jump more than 20%, helping users catch meter errors or appliance faults      early.
