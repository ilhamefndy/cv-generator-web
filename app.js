// DOM Elements
const jsonInput = document.getElementById('jsonInput');
const generateBtn = document.getElementById('generateBtn');
const loadExampleBtn = document.getElementById('loadExampleBtn');
const btnText = generateBtn.querySelector('.btn-text');
const loaderOverlay = document.getElementById('loaderOverlay');
const loaderText = document.getElementById('loaderText');
const loaderSubtext = document.getElementById('loaderSubtext');

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

// Main Generation Function
async function generateCV() {
    if (!isReady) return;
    
    const jsonData = jsonInput.value.trim();
    if (!jsonData) {
        alert("Please paste your JSON data first.");
        return;
    }
    
    // Validate JSON structure
    try {
        JSON.parse(jsonData);
    } catch(e) {
        alert("Invalid JSON data. Please check your syntax.");
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

// Event Listeners
generateBtn.addEventListener('click', generateCV);
loadExampleBtn.addEventListener('click', loadExampleData);

// Start Pyodide on load
window.addEventListener('load', initPyodide);
