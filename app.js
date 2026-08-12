// DOM Elements
const jsonInput = document.getElementById('jsonInput');
const generateBtn = document.getElementById('generateBtn');
const loadExampleBtn = document.getElementById('loadExampleBtn');
const btnText = generateBtn.querySelector('.btn-text');
const loaderOverlay = document.getElementById('loaderOverlay');
const loaderText = document.getElementById('loaderText');
const loaderSubtext = document.getElementById('loaderSubtext');

// New UI Elements
const modeToggle = document.getElementById('modeToggle');
const modeJsonLabel = document.getElementById('modeJsonLabel');
const modeRawLabel = document.getElementById('modeRawLabel');
const settingsBtn = document.getElementById('settingsBtn');
const settingsModal = document.getElementById('settingsModal');
const closeSettingsBtn = document.getElementById('closeSettingsBtn');
const saveSettingsBtn = document.getElementById('saveSettingsBtn');
const apiKeyInput = document.getElementById('apiKeyInput');
const instructionsList = document.getElementById('instructionsList');
const editorTitle = document.getElementById('editorTitle');

let pyodide = null;
let isReady = false;

// Initialize Pyodide
async function initPyodide() {
    try {
        console.log("Initializing Pyodide...");
        pyodide = await loadPyodide();
        
        btnText.textContent = "Loading Dependencies...";
        
        // Load micropip to install packages
        await pyodide.loadPackage("micropip");
        const micropip = pyodide.pyimport("micropip");
        
        // Install python-docx (and automatically lxml)
        await micropip.install("python-docx");
        
        console.log("Pyodide is ready!");
        isReady = true;
        
        // Enable generate button
        btnText.textContent = "Generate CV";
        generateBtn.disabled = false;
        
    } catch (error) {
        console.error("Failed to initialize Pyodide:", error);
        btnText.textContent = "Error loading engine";
        alert("Failed to initialize the Python engine. Please check your connection and reload.");
    }
}

// Fetch file utility
async function fetchFile(url, type = 'text') {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Failed to fetch ${url}`);
    if (type === 'arraybuffer') {
        return await response.arrayBuffer();
    }
    return await response.text();
}

// Call Gemini API
async function magicParseText(rawText) {
    const apiKey = localStorage.getItem('gemini_api_key');
    if (!apiKey) {
        alert("Please configure your Gemini API Key in AI Settings first!");
        settingsModal.classList.remove('hidden');
        throw new Error("Missing API Key");
    }

    const systemPrompt = `You are an expert CV parser. The user will provide raw unstructured text from a CV. 
You must extract all the information and output ONLY a valid JSON object matching this exact schema:
{
  "output_filename": "person_name_cv.docx",
  "header": { "full_name": "NAME", "tagline": "...", "email": "...", "phone": "...", "linkedin": "..." },
  "summary": { "text": "...", "bold_words": [] },
  "sections_order": ["education", "work", "projects", "activities", "certifications", "skills", "references"],
  "education": [{ "institution": "...", "degree": "...", "date_range": "...", "cgpa": "...", "coursework": "..." }],
  "work": [{ "company": "...", "position": "...", "date_range": "...", "bullets": ["..."] }],
  "projects": { "header_text": "PROJECTS", "entries": [{ "title": "...", "date_range": "...", "bullets": ["..."] }] },
  "certifications": { "header_text": "CERTIFICATIONS", "entries": [{ "title": "...", "date_range": "...", "bullets": ["..."] }] },
  "skills": { "header_text": "SKILLS", "categories": [{ "label": "...", "value": "..." }] },
  "references": [{ "name": "...", "title": "...", "phone": "...", "email": "..." }]
}
Do NOT wrap the response in markdown blocks (e.g. \`\`\`json). Output pure JSON. Ensure data maps correctly to these arrays/objects.`;

    const url = \`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent?key=\${apiKey}\`;
    const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            system_instruction: { parts: [{ text: systemPrompt }] },
            contents: [{ parts: [{ text: rawText }] }]
        })
    });

    if (!response.ok) {
        throw new Error("Failed to call Gemini API. Check your API key.");
    }

    const data = await response.json();
    let textOutput = data.candidates[0].content.parts[0].text;
    
    // Clean up potential markdown formatting from the AI
    textOutput = textOutput.replace(/^\`\`\`json\\s*/i, '').replace(/\\s*\`\`\`$/, '');
    
    return textOutput;
}

// Main Generation Function
async function generateCV() {
    if (!isReady) return;
    
    const inputVal = jsonInput.value.trim();
    if (!inputVal) {
        alert("Please paste your data first.");
        return;
    }
    
    let jsonData = inputVal;
    
    // If we are in Magic Text mode, parse it first!
    if (modeToggle.checked) {
        try {
            loaderOverlay.classList.remove('hidden');
            loaderText.textContent = "AI is Reading...";
            loaderSubtext.textContent = "Parsing your raw text into structured data...";
            jsonData = await magicParseText(inputVal);
        } catch(e) {
            loaderOverlay.classList.add('hidden');
            return;
        }
    }
    
    // Validate JSON structure
    try {
        JSON.parse(jsonData);
    } catch(e) {
        alert("Invalid JSON data. If using Magic Text, the AI might have failed to generate valid JSON.");
        loaderOverlay.classList.add('hidden');
        return;
    }

    try {
        // Show loader
        loaderOverlay.classList.remove('hidden');
        
        loaderText.textContent = "Downloading Template...";
        
        // Ensure directories exist in Pyodide virtual FS
        pyodide.runPython(`
            import os
            os.makedirs("/templates", exist_ok=True)
            os.makedirs("/src", exist_ok=True)
            os.makedirs("/output", exist_ok=True)
        `);
        
        // Fetch Template
        const templateBuffer = await fetchFile('./AININ_SOFEA_CV.docx', 'arraybuffer');
        pyodide.FS.writeFile('/templates/AININ_SOFEA_CV.docx', new Uint8Array(templateBuffer));
        
        // Fetch Python Source
        const sourceCode = await fetchFile('./generate_cv.py', 'text');
        pyodide.FS.writeFile('/src/generate_cv.py', sourceCode);
        
        loaderText.textContent = "Generating Document...";
        
        // Write JSON data to FS
        pyodide.FS.writeFile('/data.json', jsonData);
        
        // Run the python script
        const pythonWrapper = `
import json
import sys
import os
sys.path.append('/') # Add root to path so src.generate_cv works

# We need to monkey-patch __file__ since the original script relies on it to find the workspace
import src.generate_cv as gen_cv
# Override paths in the loaded module
from pathlib import Path
gen_cv.WORKSPACE = Path("/")
gen_cv.TEMPLATE = Path("/templates/AININ_SOFEA_CV.docx")
gen_cv.OUTPUT_DIR = Path("/output")

with open("/data.json", "r", encoding="utf-8") as f:
    data = json.loads(f.read())

out_path = gen_cv.OUTPUT_DIR / "cv_output.docx"
gen_cv.generate_cv(data, out_path)

with open(out_path, "rb") as f:
    doc_bytes = f.read()

doc_bytes
`;
        
        const docBytes = await pyodide.runPythonAsync(pythonWrapper);
        
        loaderText.textContent = "Document Ready!";
        loaderSubtext.textContent = "Starting download...";
        
        // Create Blob and Download
        const docArray = docBytes.toJs();
        const blob = new Blob([docArray], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        
        // Determine file name
        let fileName = "cv_output.docx";
        try {
            const parsed = JSON.parse(jsonData);
            if (parsed.output_filename) {
                fileName = parsed.output_filename.endsWith('.docx') ? parsed.output_filename : parsed.output_filename + '.docx';
            } else if (parsed.personal && parsed.personal.name) {
                fileName = parsed.personal.name.replace(/\\s+/g, '_') + '_CV.docx';
            }
        } catch(e) {}
        
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        
        setTimeout(() => {
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            loaderOverlay.classList.add('hidden');
        }, 1000);
        
    } catch (error) {
        console.error("Error generating CV:", error);
        alert("An error occurred while generating the CV. Check console for details.");
        loaderOverlay.classList.add('hidden');
    }
}

// Load Example Data
async function loadExampleData() {
    try {
        const exampleData = await fetchFile('./ilham_data.json', 'text');
        jsonInput.value = exampleData;
    } catch (error) {
        alert("Could not load example data.");
        console.error(error);
    }
}

// UI Toggles & Modals
modeToggle.addEventListener('change', (e) => {
    if (e.target.checked) {
        // Raw Text Mode
        modeJsonLabel.classList.remove('active');
        modeRawLabel.classList.add('active');
        jsonInput.placeholder = "Paste your raw CV text here from a PDF or Word document...\\nThe AI will organize it for you automatically!";
        instructionsList.innerHTML = `
            <li>Paste your messy, raw text into the editor.</li>
            <li>Click Generate CV.</li>
            <li>AI will structure it and generate the perfect layout!</li>
        `;
    } else {
        // JSON Mode
        modeRawLabel.classList.remove('active');
        modeJsonLabel.classList.add('active');
        jsonInput.placeholder = "Paste your JSON here...";
        instructionsList.innerHTML = `
            <li>Paste your JSON data into the editor.</li>
            <li>Ensure it matches the expected structure.</li>
            <li>Click Generate CV.</li>
        `;
    }
});

settingsBtn.addEventListener('click', () => {
    apiKeyInput.value = localStorage.getItem('gemini_api_key') || '';
    settingsModal.classList.remove('hidden');
});

closeSettingsBtn.addEventListener('click', () => {
    settingsModal.classList.add('hidden');
});

saveSettingsBtn.addEventListener('click', () => {
    const key = apiKeyInput.value.trim();
    if (key) {
        localStorage.setItem('gemini_api_key', key);
        settingsModal.classList.add('hidden');
    } else {
        localStorage.removeItem('gemini_api_key');
        settingsModal.classList.add('hidden');
    }
});

// Event Listeners
generateBtn.addEventListener('click', generateCV);
loadExampleBtn.addEventListener('click', loadExampleData);

// Start Pyodide on load
window.addEventListener('load', initPyodide);
