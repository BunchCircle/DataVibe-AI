import { GoogleGenAI, Chat } from "https://esm.run/@google/genai";
import Papa from "https://esm.run/papaparse";
import zoomPlugin from 'https://esm.run/chartjs-plugin-zoom';
import * as XLSX from "https://esm.run/xlsx";

// --- DOM Elements ---
const uploadZone = document.getElementById('upload-zone');
const fileInput = document.getElementById('file-input');
const clearDataBtn = document.getElementById('clear-data-btn');
const chatHistory = document.getElementById('chat-history');
const chatInputForm = document.getElementById('chat-input-container');
const chatInput = document.getElementById('chat-input');
const sendBtn = document.getElementById('send-btn');
const uploadContainer = document.getElementById('upload-container');
const dataSummary = document.getElementById('data-summary');
const fileNameEl = document.getElementById('file-name');
const rowCountEl = document.getElementById('row-count');
const colCountEl = document.getElementById('col-count');
const menuBtn = document.getElementById('menu-btn');
const dataSection = document.querySelector('.data-section');
const sidebarBackdrop = document.getElementById('sidebar-backdrop');
const exportOptions = document.getElementById('export-options');
const exportCsvBtn = document.getElementById('export-csv-btn');
const exportExcelBtn = document.getElementById('export-excel-btn');
const sidebarAnalysisGuide = document.getElementById('sidebar-analysis-guide'); 
const filterSection = document.getElementById('filter-section');
const filterList = document.getElementById('filter-list');
const clearFiltersBtn = document.getElementById('clear-filters-btn');
const stagingSection = document.getElementById('staging-section');
const stagedFiltersList = document.getElementById('staged-filters-list');
const applyFiltersBtn = document.getElementById('apply-filters-btn');
const cancelStagingBtn = document.getElementById('cancel-staging-btn');
const welcomeScreen = document.getElementById('welcome-screen');
const welcomeUploadPrompt = document.getElementById('welcome-upload-prompt');
// Modal Elements
const howToUseBtn = document.getElementById('how-to-use-btn');
const helpModal = document.getElementById('help-modal');
const closeHelpBtn = document.getElementById('close-help-btn');


// --- App State ---
let ai;
let chat = null; // The stateful chat session with the AI
let originalData = null; // The master, unfiltered data
let activeData = null; // The data currently being displayed/analyzed (can be filtered)
let activeFilters = []; // { column, value }
let stagedFilters = []; // Filters selected by user, waiting to be applied
let isFilteredState = false; // Is the app currently showing a filtered subset of data?
const chartInstances = new Map();
const MAX_FILE_SIZE = 49 * 1024 * 1024; // 49MB

// --- Chart Interaction State ---
let clickTimer = null;
const CLICK_DELAY = 250; // ms threshold for double-click
let currentTooltip = null;


// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
    Chart.register(zoomPlugin);
    initializeAi();
    setupEventListeners();
    observeChatInputResize();
});

function setupEventListeners() {
    uploadZone.addEventListener('dragover', handleDragOver);
    uploadZone.addEventListener('dragleave', handleDragLeave);
    uploadZone.addEventListener('drop', handleDrop);
    fileInput.addEventListener('change', handleFileSelect);
    clearDataBtn.addEventListener('click', clearData);
    chatInputForm.addEventListener('submit', handleSendMessage);
    chatInput.addEventListener('input', autoResizeTextarea);
    chatInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            chatInputForm.requestSubmit();
        }
    });
    menuBtn.addEventListener('click', toggleSidebar);
    sidebarBackdrop.addEventListener('click', toggleSidebar);
    exportCsvBtn.addEventListener('click', exportAsCsv);
    exportExcelBtn.addEventListener('click', exportAsExcel);
    clearFiltersBtn.addEventListener('click', handleClearFilters);
    applyFiltersBtn.addEventListener('click', handleApplyStagedFilters);
    cancelStagingBtn.addEventListener('click', handleCancelStaging);
    welcomeUploadPrompt.addEventListener('click', () => fileInput.click());
    
    // Modal Listeners
    howToUseBtn.addEventListener('click', () => {
        helpModal.classList.add('visible');
    });
    closeHelpBtn.addEventListener('click', () => {
        helpModal.classList.remove('visible');
    });
    helpModal.addEventListener('click', (e) => {
        if (e.target === helpModal) {
            helpModal.classList.remove('visible');
        }
    });
}

function toggleSidebar() {
    dataSection.classList.toggle('open');
    sidebarBackdrop.classList.toggle('visible');
}

function initializeAi() {
    try {
        ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
    } catch (error) {
        addErrorMessageToChat('Initialization Failed', 'Could not connect to the AI service. Please check your API key and network connection.');
    }
}

// --- Data Cleaning & Inference ---

/**
 * Automatically cleans "dirty" numeric columns.
 * Specifically targets columns that act like numbers but contain non-numeric characters
 * like currency symbols (₹, $, etc.), commas, or spaces.
 * 
 * Mutates the data array in place.
 */
function cleanAndNormalizeData(data, fields) {
    if (!data || data.length === 0) return;

    const SAMPLE_LIMIT = 100;
    const CLEANING_THRESHOLD = 0.8; // 80% confidence required to alter a column
    const sample = data.slice(0, SAMPLE_LIMIT);
    const fieldsToClean = [];

    fields.forEach(field => {
        let numericLikeCount = 0;
        let validValueCount = 0;
        let needsCleaning = false;

        for (const row of sample) {
            const val = row[field];
            if (val === null || val === undefined || String(val).trim() === '') continue;
            validValueCount++;

            if (typeof val === 'number') {
                numericLikeCount++;
            } else if (typeof val === 'string') {
                // 1. Check if it's already a clean valid number string
                if (!isNaN(Number(val)) && val.trim() !== '') {
                    numericLikeCount++;
                    continue;
                }

                // 2. Check if it becomes a number after stripping "noise"
                // Remove everything that is NOT a digit, dot (.), or minus sign (-)
                const cleaned = val.replace(/[^0-9.-]/g, '');
                
                // Ensure the cleaned string is a valid number and actually contains digits
                if (cleaned !== '' && !isNaN(Number(cleaned)) && /\d/.test(cleaned)) {
                    // Check if specific "bad" chars existed in the original string (to verify it needed cleaning)
                    // e.g., "₹ 150" has '₹' and ' '. "150" does not.
                    if (/[^0-9.-]/.test(val)) {
                        needsCleaning = true;
                    }
                    numericLikeCount++;
                }
            }
        }

        // If most values look like numbers (or can be made into numbers), and we detected some dirt, mark for cleaning.
        if (validValueCount > 0 && (numericLikeCount / validValueCount) >= CLEANING_THRESHOLD && needsCleaning) {
            fieldsToClean.push(field);
        }
    });

    if (fieldsToClean.length > 0) {
        console.log(`Auto-cleaning columns: ${fieldsToClean.join(', ')}`);
        
        // Apply cleaning to the entire dataset
        data.forEach(row => {
            fieldsToClean.forEach(field => {
                const val = row[field];
                if (typeof val === 'string') {
                    const cleaned = val.replace(/[^0-9.-]/g, '');
                    if (cleaned !== '' && !isNaN(Number(cleaned))) {
                        // Mutate the row value to be a pure Number
                        row[field] = Number(cleaned);
                    }
                }
            });
        });
    }
}

const MAX_SAMPLE_SIZE = 100;
const CATEGORICAL_UNIQUENESS_THRESHOLD = 0.5; // If unique values are less than 50% of sample size
const CATEGORICAL_MAX_UNIQUE_VALUES = 50; // And unique values are less than 50

/**
 * Infers the data type of each column based on a sample of the data.
 * @param {Array<Object>} data - The array of data rows.
 * @param {Array<string>} fields - The list of column names.
 * @returns {Object} A map of column names to their inferred types (e.g., 'numerical', 'temporal', 'categorical', 'string').
 */
function inferColumnTypes(data, fields) {
    const types = {};
    const sample = data.slice(0, MAX_SAMPLE_SIZE);
    if (sample.length === 0) return types;

    for (const field of fields) {
        let isNumeric = true;
        let isTemporal = true;
        const uniqueValues = new Set();

        for (const row of sample) {
            const value = row[field];
            if (value === null || value === undefined || value === '') {
                continue; // Skip empty values for inference
            }

            uniqueValues.add(value);

            // Check for numeric
            if (isNumeric) {
                const cleanedValue = String(value).replace(/[\$,%]/g, '').trim();
                if (cleanedValue === '' || isNaN(Number(cleanedValue))) {
                    isNumeric = false;
                }
            }
            
            // Check for temporal
            if (isTemporal) {
                // Exclude raw numbers from being parsed as dates (e.g., year '2023')
                if (typeof value === 'number' || isNaN(Date.parse(String(value)))) {
                     isTemporal = false;
                }
            }
        }
        
        if (isNumeric) {
            types[field] = 'numerical';
        } else if (isTemporal) {
            types[field] = 'temporal';
        } else if (uniqueValues.size / sample.length < CATEGORICAL_UNIQUENESS_THRESHOLD && uniqueValues.size <= CATEGORICAL_MAX_UNIQUE_VALUES) {
            types[field] = 'categorical';
        } else {
            types[field] = 'string';
        }
    }
    return types;
}


// --- File Handling ---
function handleDragOver(e) {
    e.preventDefault();
    uploadZone.classList.add('drag-over');
}

function handleDragLeave(e) {
    e.preventDefault();
    uploadZone.classList.remove('drag-over');
}

function handleDrop(e) {
    e.preventDefault();
    uploadZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    processFile(file);
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    processFile(file);
}

function handleSuccessfulParse(results, fileName) {
    // 1. Auto-Clean Data (Remove currency symbols, etc.)
    cleanAndNormalizeData(results.data, results.meta.fields);

    // 2. Infer data types for smarter AI analysis
    const inferredTypes = inferColumnTypes(results.data, results.meta.fields);
    results.meta.inferredTypes = inferredTypes;
    
    originalData = results;
    activeData = results; // Initially, active data is the same as original
    addPreviewToChat(activeData);
    updateUiOnDataLoad(fileName, activeData);
    initializeChatSession(results);
    generateInitialDataSummary(results);
}

function processFile(file) {
    if (!file) return;

    if (file.size > MAX_FILE_SIZE) {
        addErrorMessageToChat('File Too Large', 'The maximum file size is 49MB. Please upload a smaller file.');
        return;
    }

    const isCsv = file.type.match('text/csv') || file.name.endsWith('.csv');
    const isExcel = file.type.match(/spreadsheetml|ms-excel/) || file.name.endsWith('.xlsx') || file.name.endsWith('.xls');

    if (isCsv) {
        Papa.parse(file, {
            header: true,
            skipEmptyLines: true,
            complete: (results) => {
                handleSuccessfulParse(results, file.name);
            },
            error: (error) => {
                addErrorMessageToChat('CSV Parsing Error', error.message);
            }
        });
    } else if (isExcel) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);

                if (jsonData.length === 0) {
                     addErrorMessageToChat('Empty File', 'The Excel file appears to be empty or the first sheet has no data.');
                     return;
                }

                const fields = Object.keys(jsonData[0]);
                const results = {
                    data: jsonData,
                    meta: { fields: fields }
                };
                handleSuccessfulParse(results, file.name);

            } catch (error) {
                addErrorMessageToChat('Excel Parsing Error', 'Could not read the Excel file. It might be corrupted or in an unsupported format.');
                console.error(error);
            }
        };
        reader.onerror = (error) => {
             addErrorMessageToChat('File Read Error', 'An unexpected error occurred while trying to read the file.');
             console.error("FileReader error:", error);
        };
        reader.readAsArrayBuffer(file);
    } else {
        addErrorMessageToChat('Unsupported File Type', 'Please upload a CSV, XLSX, or XLS file.');
    }
}

function addPreviewToChat(dataObject) {
    const { data, meta } = dataObject;
    const rowCount = data.length;
    // Truncate column list for display to prevent massive overflow
    let columnNames = meta.fields.join(', ');
    if (columnNames.length > 150) {
        columnNames = columnNames.substring(0, 150) + '...';
    }

    // Create table header
    let headerHtml = '<tr>';
    meta.fields.forEach(field => {
        headerHtml += `<th>${field}</th>`;
    });
    headerHtml += '</tr>';

    // Create table body rows
    let bodyHtml = '';
    const previewRows = data.slice(0, 5);
    previewRows.forEach(row => {
        bodyHtml += '<tr>';
        meta.fields.forEach(field => {
            bodyHtml += `<td>${row[field] || ''}</td>`;
        });
        bodyHtml += '</tr>';
    });

    const previewHtml = `
        <div class="preview-in-chat-metadata">
            <p><strong>Rows:</strong> ${rowCount}</p>
            <p><strong>Columns:</strong> ${columnNames}</p>
        </div>
        <div class="preview-in-chat-table-wrapper">
            <table class="preview-in-chat-table">
                <thead>${headerHtml}</thead>
                <tbody>${bodyHtml}</tbody>
            </table>
        </div>
    `;
    
    const messageEl = document.createElement('div');
    messageEl.classList.add('message', 'ai-message', 'preview-in-chat');
    messageEl.innerHTML = previewHtml;

    if (isFilteredState) {
        messageEl.dataset.filterMessage = 'true';
    }

    chatHistory.appendChild(messageEl);
}


function updateUiOnDataLoad(fileName, dataObject) {
    // Hide welcome screen
    welcomeScreen.classList.add('hidden');

    // Enable chat
    chatInput.disabled = false;
    sendBtn.disabled = false;

    // Update sidebar UI
    uploadContainer.classList.add('hidden');
    dataSummary.classList.remove('hidden');
    // Hide the persistent guide when data is loaded
    sidebarAnalysisGuide.classList.add('hidden');
    exportOptions.classList.remove('hidden');
    clearDataBtn.classList.remove('hidden');
    filterSection.classList.add('hidden');
    stagingSection.classList.add('hidden');


    fileNameEl.textContent = fileName;
    updateSidebarStats(dataObject);
}

function updateSidebarStats(dataObject) {
    rowCountEl.textContent = dataObject.data.length;
    colCountEl.textContent = dataObject.meta.fields.length;
}

function clearData() {
    originalData = null;
    activeData = null;
    activeFilters = [];
    stagedFilters = [];
    isFilteredState = false;
    chat = null; // Clear the chat session
    fileInput.value = ''; // Reset file input
    chatInput.disabled = true;
    sendBtn.disabled = true;

    // Clear all chart instances and remove them from the DOM
    chartInstances.forEach(instance => instance.destroy());
    chartInstances.clear();
    
    // Clear chat history by removing all generated messages
    chatHistory.querySelectorAll('.message, .chart-message, .preview-in-chat, .error-message').forEach(el => el.remove());

    // Show the welcome screen
    welcomeScreen.classList.remove('hidden');

    // Reset sidebar UI
    dataSummary.classList.add('hidden');
    // Show the persistent guide when data is cleared
    sidebarAnalysisGuide.classList.remove('hidden');
    uploadContainer.classList.remove('hidden');
    exportOptions.classList.add('hidden');
    clearDataBtn.classList.add('hidden');
    filterSection.classList.add('hidden');
    stagingSection.classList.add('hidden');

    if (dataSection.classList.contains('open')) {
        toggleSidebar();
    }
}

// --- Data Filtering ---
function applyFilters() {
    if (activeFilters.length > 0) {
        let filteredRows = originalData.data;
        activeFilters.forEach(filter => {
            const { column, value } = filter;
            filteredRows = filteredRows.filter(row => String(row[column]) === String(value));
        });

        activeData = {
            ...originalData,
            data: filteredRows,
        };
        isFilteredState = true;
        updateUiForFilters();
    } else {
        // Clearing the filter
        activeData = originalData;
        isFilteredState = false;
        filterSection.classList.add('hidden');
        updateSidebarStats(activeData);
    }
}

function handleClearFilters() {
    // 1. Select all messages that were generated while filters were active.
    const messagesToRemove = chatHistory.querySelectorAll('[data-filter-message="true"]');
    
    // 2. Animate them out and then remove them from the DOM.
    messagesToRemove.forEach(msg => {
        msg.classList.add('fade-out-and-remove');
        // The 'animationend' event ensures the element is removed only after the animation finishes.
        msg.addEventListener('animationend', () => {
            msg.remove();
        }, { once: true });
    });
    
    // 3. Reset the filter state.
    activeFilters = [];
    applyFilters(); // This will reset activeData and set isFilteredState to false.
    
    // 4. Add a confirmation message.
    addMessageToChat('All filters have been cleared. Showing all data.', 'ai');
}


function updateUiForFilters() {
    filterList.innerHTML = '';
    activeFilters.forEach(filter => {
        const item = document.createElement('div');
        item.classList.add('filter-item');
        item.innerHTML = `<span>${filter.column}</span> = <span>${filter.value}</span>`;
        filterList.appendChild(item);
    });

    filterSection.classList.remove('hidden');
    updateSidebarStats(activeData);

    const filterSummary = activeFilters.map(f => `<em>${f.column}</em> is <em>${f.value}</em>`).join(' AND ');
    addMessageToChat(`<strong>Filters applied:</strong> Showing data where ${filterSummary}. All analyses will now use this subset.`, 'ai');
    addPreviewToChat(activeData);
    generateAndDisplayFollowUpQuestions('<strong>What\'s next?</strong> Here are some ideas for the filtered data:');
}

function updateStagingUi() {
    if (stagedFilters.length > 0) {
        stagedFiltersList.innerHTML = '';
        stagedFilters.forEach((filter, index) => {
            const item = document.createElement('div');
            item.classList.add('staged-filter-item');
            item.innerHTML = `
                <span><strong>${filter.column}:</strong> ${filter.value}</span>
                <button class="remove-staged-btn" data-index="${index}" aria-label="Remove filter">&times;</button>
            `;
            stagedFiltersList.appendChild(item);
        });
        
        // Add event listeners to new remove buttons
        stagedFiltersList.querySelectorAll('.remove-staged-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const indexToRemove = parseInt(e.currentTarget.dataset.index, 10);
                stagedFilters.splice(indexToRemove, 1);
                updateStagingUi(); // Re-render
            });
        });

        stagingSection.classList.remove('hidden');
    } else {
        stagingSection.classList.add('hidden');
    }
}

function handleApplyStagedFilters() {
    activeFilters = [...activeFilters, ...stagedFilters];
    // Remove duplicates
    activeFilters = activeFilters.filter((filter, index, self) =>
        index === self.findIndex((f) => (
            f.column === filter.column && f.value === filter.value
        ))
    );
    stagedFilters = [];
    updateStagingUi();
    applyFilters();
}

function handleCancelStaging() {
    stagedFilters = [];
    updateStagingUi();
}

// --- Chat & AI ---

function createSystemInstruction(dataObject) {
    const { data, meta } = dataObject;
    
    const columnsWithTypes = meta.fields.map(field => 
        `${field} (${meta.inferredTypes[field] || 'string'})`
    ).join(', ');

    const sampleDataForContext = Papa.unparse(data.slice(0, 20));

    return `
You are an expert data analyst AI inside a web application. Your task is to analyze a dataset based on a user's question and provide concise answers, along with visualizations. You will maintain a conversation history.

The user has uploaded a dataset with the following schema and data sample:
- Columns with inferred data types: ${columnsWithTypes}
- Sample Data (first 20 rows in CSV format):
${sampleDataForContext}

--- CHARTING & AGGREGATION LOGIC ---
When the user asks for a visualization, you MUST use the inferred data types to make intelligent choices about the chart type and aggregation method.
- **Temporal vs. Numerical**: Use 'line' or 'area' charts to show trends over time.
- **Categorical vs. Numerical**: Use 'bar', 'pie', or 'donut' charts. For bar charts, sort by the numerical value descending.
- **Numerical vs. Numerical**: Use 'scatter' plots to show relationships.
- **Aggregation**:
  - For 'numerical' value columns, default to 'sum' or 'average'.
  - For 'categorical' columns that are being measured, use 'count'.

--- RESPONSE RULES ---
Always provide your response in two parts, separated by "---VIZ---".

PART 1: Answer
A very concise, text-based answer to the user's question, styled with simple HTML tags (e.g., <strong>, <p>). Keep the text to a minimum. If the user's query is a follow-up, your text should acknowledge the context.

PART 2: Visualization
This part contains a single JSON object.

--- JSON FORMAT & LOGIC ---
First, determine if the user's request is a "simple" or "complex" analysis.
- A "simple" analysis is a direct aggregation (sum, average, count) on one of the existing columns. Example: "What is the total sales for each product_category?".
- A "complex" analysis requires data manipulation *before* aggregation, such as grouping data into time bins (e.g., "every two hours"), creating new categories based on conditions, or other advanced transformations. Example: "Visualize sales on a two-hourly basis throughout the day."

Based on your determination, your JSON response MUST be in one of the following two formats:

**FORMAT A: For Simple Analysis**
{
  "analysisType": "simple",
  "chartConfig": {
    "chartType": "bar" | "line" | "area" | "pie" | "donut" | "scatter",
    "title": "A descriptive chart title",
    "aggregation": "sum" | "average" | "count",
    "xAxisColumn": "column_name_for_x_axis",
    "yAxisColumn": "column_name_for_y_axis",
    "categoryColumn": "column_name_for_pie_or_donut_labels",
    "valueColumn": "column_name_for_pie_or_donut_values",
    "explanation": "A concise, bulleted HTML list (e.g., <ul><li>Insight 1 with numbers.</li><li>Insight 2.</li></ul>) summarizing the chart's key findings."
  }
}
- Use this format when the frontend can calculate the data itself using the specified aggregation and columns.
- **CRITICAL**: Column names MUST exist in the schema.
- Aggregation Logic: For "count", omit yAxisColumn/valueColumn. For "sum" or "average", both label and value columns are required.
- For scatter charts, omit \`aggregation\` and provide \`xAxisColumn\` and \`yAxisColumn\`.

**FORMAT B: For Complex Analysis**
{
  "analysisType": "complex",
  "chartConfig": {
    "chartType": "bar" | "line" | "area",
    "title": "A descriptive chart title",
    "explanation": "A concise, bulleted HTML list...",
    "axisTitles": {
        "x": "X-Axis Title (e.g., 'Time of Day')",
        "y": "Y-Axis Title (e.g., 'Total Sales')"
    }
  },
  "chartData": {
    "labels": ["Label 1", "Label 2", "Label 3"],
    "values": [150, 230, 180]
  }
}
- Use this format when the request requires custom data processing that the frontend cannot do (like time binning).
- You MUST perform the aggregation yourself based on the full dataset context and return the final data points in \`chartData\`.
- \`chartData.labels\` and \`chartData.values\` MUST be arrays of the same length.
- In \`chartConfig\` for complex analysis, you do not need aggregation or column keys. Provide \`axisTitles\` instead.

--- SCENARIOS ---
If the user's question is vague (e.g., "visualize this data"), you MUST generate 2-3 different visualizations.
- Your text answer should be a brief introduction.
- The visualization part must be a single JSON object with a "visualizations" key. The value should be an array of 2-3 objects, where each object follows either "FORMAT A" or "FORMAT B" from above. Example:
{
  "visualizations": [
    { "analysisType": "simple", "chartConfig": { /* ... */ } },
    { "analysisType": "simple", "chartConfig": { /* ... */ } }
  ]
}
`;
}

function initializeChatSession(dataObject) {
    if (!ai) {
        addErrorMessageToChat('AI Not Initialized', 'Cannot start chat session.');
        return;
    }
    
    const systemInstruction = createSystemInstruction(dataObject);

    chat = ai.chats.create({
        model: 'gemini-2.5-flash',
        config: {
            systemInstruction: systemInstruction,
        },
    });
}

async function generateInitialDataSummary(dataObject) {
    if (!chat) {
        addErrorMessageToChat('Chat Not Initialized', 'Cannot generate summary because the chat session is not available.');
        return;
    }
    
    showLoadingIndicator();

    try {
        const { data } = dataObject;
        const sampleData = Papa.unparse(data.slice(0, 5));

        const prompt = `
A new dataset has just been uploaded. Your task is to provide an immediate, insightful overview with key visualizations to get the user started.

Here's the sample data (first 5 rows) for initial context:
${sampleData}

--- TASK ---
Provide a response that follows the two-part format (text and viz) you were instructed on.
1.  **Text Part**: A single welcoming <p> tag that briefly describes the dataset's likely content.
2.  **Viz Part**: A JSON object with a "visualizations" key, containing an array of 2 to 3 different chart configurations that provide a broad overview (e.g., distributions, totals), making sure to use your knowledge of the column data types to pick appropriate charts.
`;
        const response = await chat.sendMessage({ message: prompt });
        
        await processAiResponse(response.text);

    } catch (error) {
        addErrorMessageToChat('Initial Analysis Failed', 'I encountered an issue while creating the initial summary and charts. Please try uploading the file again.');
        console.error(error);
    } finally {
        removeLoadingIndicator();
    }
}

async function generateAndDisplayFollowUpQuestions(title) {
    if (!ai || !activeData) return;

    try {
        const { meta } = activeData;
        const columnsWithTypes = meta.fields.map(field =>
            `${field} (${meta.inferredTypes[field] || 'string'})`
        ).join(', ');

        const prompt = `
You are a helpful data analyst assistant. Your task is to suggest insightful follow-up questions based on a dataset's schema and inferred data types.

Dataset Schema:
[${columnsWithTypes}]

--- TASK ---
Generate a JSON array containing 2 to 3 insightful yet simple and direct analysis questions a user could ask about this data.

--- GUIDELINES ---
- Use the data types to ask appropriate questions. For example, suggest line charts for 'temporal' data, bar charts for 'categorical', and scatter plots for two 'numerical' columns.
- Questions must be diverse and explore different aspects of the data.
- Good examples: "Show the trend of [Numerical Column] over [Temporal Column].", "What is the total [Numerical Column] for each [Categorical Column]?", "Is there a relationship between [Numerical Column 1] and [Numerical Column 2]?"
- AVOID complex, multi-part questions.

--- RESPONSE FORMAT ---
- Respond with ONLY the JSON array of strings.
- Example: ["What is the distribution of sales by region?", "Show the top 5 selling products by total revenue."]
- DO NOT use markdown like \`\`\`json.`;

        // Use a stateless call for this helper function to not pollute the main chat history
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
        });

        const jsonString = response.text.trim().replace(/```json/g, '').replace(/```/g, '');
        const suggestedQuestions = JSON.parse(jsonString);

        if (Array.isArray(suggestedQuestions) && suggestedQuestions.length > 0) {
            addSuggestedQuestionsToChat(suggestedQuestions, title);
        }

    } catch (error) {
        console.error("Could not generate follow-up questions:", error);
        // Fail silently, this is a non-essential feature.
    }
}


function addSuggestedQuestionsToChat(questions, titleText) {
    const containerEl = document.createElement('div');
    containerEl.classList.add('message', 'ai-message');
    
    if (isFilteredState) {
        containerEl.dataset.filterMessage = 'true';
    }

    const title = document.createElement('p');
    title.innerHTML = titleText;
    containerEl.appendChild(title);

    const buttonGroup = document.createElement('div');
    buttonGroup.classList.add('suggested-questions-buttons');
    questions.slice(0, 4).forEach(q => { // Max 4 questions
        const button = document.createElement('button');
        button.textContent = q;
        button.classList.add('suggested-question-btn');
        button.addEventListener('click', async () => {
            // Disable all buttons in this group after one is clicked
            containerEl.querySelectorAll('.suggested-question-btn').forEach(btn => {
                btn.disabled = true;
            });
            
            // Manually trigger the send message flow
            chatInput.value = q;
            // The handleSendMessage function will add the message to the chat.
            handleSendMessage({ preventDefault: () => {} });
        });
        buttonGroup.appendChild(button);
    });

    containerEl.appendChild(buttonGroup);
    chatHistory.appendChild(containerEl);
    chatHistory.scrollTop = chatHistory.scrollHeight;
}

async function handleSendMessage(e) {
    e.preventDefault();
    const userQuery = chatInput.value.trim();
    if (!userQuery || !activeData) return;

    addMessageToChat(userQuery, 'user');
    chatInput.value = '';
    showLoadingIndicator();

    try {
        await classifyUserIntent(userQuery);
    } catch (error) {
        addErrorMessageToChat('Request Failed', 'I was unable to process your request. Please try rephrasing your question.');
        console.error("Main error handler:", error);
        removeLoadingIndicator();
    }
}

async function classifyUserIntent(userQuery) {
    if (!ai) {
        throw new Error("AI is not initialized.");
    }
    const prompt = `
You are an intent classifier. Your task is to categorize a user's request into one of three types: 'ANALYSIS', 'TRANSFORMATION', or 'COMPLEX_ANALYSIS'.

- 'TRANSFORMATION': The user wants to permanently change the dataset. This includes removing rows, creating new columns, or renaming columns.
  Examples: "remove all rows where sales are 0", "create a profit column from sales and cost", "rename 'cust_id' to 'CustomerID'".

- 'ANALYSIS': The user is asking a direct question that can be answered with a single aggregation or visualization from the existing data.
  Examples: "what are the total sales by region?", "show me a chart of sales over time", "count the number of products".

- 'COMPLEX_ANALYSIS': The user's request requires multiple steps to answer. This often involves creating temporary calculations, filtering, sorting, and then visualizing the result. These steps are done for a single query and DO NOT permanently change the data.
  Examples: "show me the top 5 products by profit margin", "compare the monthly sales growth for the last two quarters", "which category had the highest sales in the second half of the year?".

User request: "${userQuery}"

Respond with only one word: 'ANALYSIS', 'TRANSFORMATION', or 'COMPLEX_ANALYSIS'.`;

    const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: prompt,
    });
    
    const intent = response.text.trim().toUpperCase();

    if (intent === 'TRANSFORMATION') {
        await handleTransformation(userQuery);
    } else if (intent === 'COMPLEX_ANALYSIS') {
        await handleComplexAnalysis(userQuery);
    } else {
        await handleAnalysis(userQuery);
    }
}

async function handleAnalysis(userQuery) {
    if (!chat) {
        addErrorMessageToChat('Chat Not Initialized', 'Cannot analyze because the chat session is not available.');
        removeLoadingIndicator();
        return;
    }
    
    let fullQuery = userQuery;
    if (activeFilters.length > 0) {
        const filterContext = activeFilters.map(f => `${f.column} is '${f.value}'`).join(' AND ');
        fullQuery = `The current dataset is filtered where ${filterContext}. Now, please answer this question based on the filtered data: "${userQuery}"`;
    }

    try {
        const response = await chat.sendMessage({ message: fullQuery });
        await processAiResponse(response.text);
    } catch(e) {
        addErrorMessageToChat('Analysis Failed', 'There was a problem analyzing your data. The AI model may be temporarily unavailable.');
        console.error(e);
    } finally {
        removeLoadingIndicator();
    }
}

function buildPlannerPrompt(userQuery) {
    const { meta } = activeData;
    const columnsWithTypes = meta.fields.map(field =>
        `${field} (${meta.inferredTypes[field] || 'string'})`
    ).join(', ');

    return `
You are a data analysis planner. Your job is to break down a complex user query into a sequence of simple, executable steps. The analysis is temporary and does not modify the original dataset.

Dataset Schema (Available Columns and their inferred types):
- ${columnsWithTypes}

User Request:
"${userQuery}"

--- SUPPORTED ACTIONS & JSON FORMAT ---
You must generate a JSON array of step objects. Each step must have an "action", "explanation", and "params".

1.  **create_column**:
    {
      "action": "create_column",
      "explanation": "First, I'll calculate the [new column name]...",
      "params": {
        "newColumn": "new_column_name",
        "column1": "first_operand_column",
        "operator": "add" | "subtract" | "multiply" | "divide",
        "column2": "second_operand_column"
      }
    }

2.  **remove_rows**:
    {
      "action": "remove_rows",
      "explanation": "Then, I'll filter the data to only include rows where...",
      "params": {
        "conditions": [{
          "column": "column_name",
          "operator": "equals" | "not_equals" | "greater_than" | "less_than" | "contains",
          "value": "some_value"
        }]
      }
    }

3.  **aggregate**:
    {
      "action": "aggregate",
      "explanation": "Next, I will group by [column] and calculate the [aggregation]...",
      "params": {
        "groupBy": "column_to_group_by",
        "aggregation": "sum" | "average" | "count",
        "valueColumn": "column_to_aggregate" // Omit for 'count'
      }
    }
    - This action transforms the data. The new aggregated value column will be named 'aggregation_valueColumn' (e.g., 'average_Profit'). The groupBy column keeps its name.

4.  **sort**:
    {
      "action": "sort",
      "explanation": "I'll sort the results by [column]...",
      "params": {
        "column": "column_to_sort_by",
        "order": "ascending" | "descending"
      }
    }

5.  **limit**:
    {
      "action": "limit",
      "explanation": "Then, I'll take the top [number] results...",
      "params": { "count": 5 }
    }

6.  **visualize** (This MUST be the final step):
    {
      "action": "visualize",
      "explanation": "Finally, here is a chart showing the result.",
      "params": {
        "chartType": "bar" | "line" | "pie" | "donut",
        "title": "A descriptive chart title",
        "xAxisColumn": "column_for_x_axis_or_labels",
        "yAxisColumn": "column_for_y_axis_or_values"
      }
    }

--- RESPONSE RULES ---
- Respond with ONLY the JSON array of steps.
- Do not use markdown like \`\`\`json.
- Ensure all column names in the JSON exactly match the schema.
- The plan must end with a "visualize" action.
- If the request cannot be fulfilled, respond with: [{ "action": "error", "explanation": "I'm sorry, I can't fulfill that request with the available tools." }]`;
}


async function handleComplexAnalysis(userQuery) {
    try {
        addMessageToChat("That's an interesting question. Let me break that down and analyze it for you...", 'ai');
        const prompt = buildPlannerPrompt(userQuery);

        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
        });
        
        const jsonString = response.text.trim().replace(/```json/g, '').replace(/```/g, '');
        const plan = JSON.parse(jsonString);

        if (plan.length > 0 && plan[0].action === 'error') {
            addErrorMessageToChat('Analysis Not Possible', plan[0].explanation);
            return;
        }

        await executeAnalysisPlan(plan);

    } catch (error) {
        addErrorMessageToChat('Complex Analysis Failed', 'I was unable to create or execute a plan for your request. Please try rephrasing it.');
        console.error("Complex analysis error:", error);
    } finally {
        removeLoadingIndicator();
    }
}

async function executeAnalysisPlan(plan) {
    // Create a deep copy of the active data to work with, leaving the original untouched.
    let tempData = JSON.parse(JSON.stringify(activeData));

    for (const step of plan) {
        // Announce the step the AI is taking.
        addMessageToChat(`<em>Step ${plan.indexOf(step) + 1}:</em> ${step.explanation}`, 'ai');

        if (step.action === 'visualize') {
            const chartConfig = step.params;
            const labels = tempData.data.map(row => row[chartConfig.xAxisColumn]);
            const values = tempData.data.map(row => row[chartConfig.yAxisColumn]);

            const precomputedData = { labels, values };
            const axisTitles = { x: chartConfig.xAxisColumn, y: chartConfig.yAxisColumn };
            
            const finalChartConfig = {
                ...chartConfig,
                axisTitles: axisTitles,
                explanation: step.explanation
            };
            
            await addChartMessageToChat(finalChartConfig, precomputedData);
            await generateAndDisplayFollowUpQuestions('<strong>What\'s next?</strong> Here are some ideas:');
            return; // Plan finished
        }

        // For all non-visualize steps, transform the data.
        tempData = executeDataStep(tempData, step);
    }
}

function executeDataStep(dataObject, step) {
    const params = step.params;
    switch (step.action) {
        case 'create_column':
            return transformCreateColumn(dataObject, params);
        case 'remove_rows':
            return transformRemoveRows(dataObject, params.conditions);
        case 'rename_column':
            return transformRenameColumn(dataObject, params);
        case 'aggregate':
            return executeAggregate(dataObject, params);
        case 'sort':
            return executeSort(dataObject, params);
        case 'limit':
            return executeLimit(dataObject, params);
        default:
            throw new Error(`Unsupported plan action: ${step.action}`);
    }
}

function buildTransformationPrompt(userQuery) {
    const { meta } = originalData; // Always use originalData for schema
    const columnsWithTypes = meta.fields.map(field =>
        `${field} (${meta.inferredTypes[field] || 'string'})`
    ).join(', ');

    return `
You are a data transformation expert. Your task is to convert a user's natural language request into a structured JSON command.

Dataset Schema:
- Columns with inferred types: ${columnsWithTypes}

User Request:
"${userQuery}"

--- SUPPORTED ACTIONS & JSON FORMAT ---

1.  **Remove Rows**:
    {
      "action": "remove_rows",
      "explanation": "A short sentence explaining what was done.",
      "conditions": [{
        "column": "column_name",
        "operator": "equals" | "not_equals" | "greater_than" | "less_than" | "contains",
        "value": "some_value"
      }]
    }
    - You can have multiple conditions. All conditions must be met (AND logic).
    - For numeric comparisons, ensure the 'value' is a number, not a string.

2.  **Create Column**:
    {
      "action": "create_column",
      "explanation": "A short sentence explaining what was done.",
      "newColumn": "new_column_name",
      "column1": "first_operand_column",
      "operator": "add" | "subtract" | "multiply" | "divide",
      "column2": "second_operand_column"
    }

3.  **Rename Column**:
    {
        "action": "rename_column",
        "explanation": "A short sentence explaining what was done.",
        "oldColumn": "current_column_name",
        "newColumn": "new_column_name"
    }

--- RESPONSE RULES ---
- Respond with ONLY the JSON configuration.
- Do not use markdown like \`\`\`json.
- Ensure all column names in the JSON exactly match the schema. If a close match is found, use the correct one from the schema.
- If the request is ambiguous or cannot be translated into one of the supported actions, respond with: { "action": "error", "explanation": "Your request is unclear or not supported. Please try rephrasing." }`;
}

async function handleTransformation(userQuery) {
    try {
        const prompt = buildTransformationPrompt(userQuery);

        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
        });
        
        const jsonString = response.text.trim().replace(/```json/g, '').replace(/```/g, '');
        const config = JSON.parse(jsonString);

        if (config.action === 'error') {
            addErrorMessageToChat('Transformation Not Applied', config.explanation);
            return;
        }
        
        // Always apply transformations to the master originalData
        const result = applyTransformation(config, originalData);
        
        if (result.success) {
            // Update both original and active data
            originalData = result.newData;
            activeData = result.newData;
            
            // Clear any active filter as it might be invalid now
            if(activeFilters.length > 0) {
                activeFilters = [];
                isFilteredState = false;
                filterSection.classList.add('hidden');
            }

            // The data schema may have changed, so we MUST restart the chat session
            // with the new context.
            initializeChatSession(originalData);
            
            addMessageToChat(result.message, 'ai');
            updateSidebarStats(originalData);
            addPreviewToChat(originalData);
            addMessageToChat("The data has been transformed. The analysis context has been updated with the new schema.", 'ai');
            await generateAndDisplayFollowUpQuestions('<strong>What\'s next?</strong> Here are some ideas for the transformed data:');
        } else {
             addErrorMessageToChat('Transformation Failed', result.message);
        }

    } catch (error) {
        addErrorMessageToChat('Transformation Failed', 'The requested data transformation could not be applied. Please check your query and try again.');
        console.error("Transformation error:", error);
    } finally {
        removeLoadingIndicator();
    }
}

function applyTransformation(config, sourceData) {
    let newDataObject = JSON.parse(JSON.stringify(sourceData));
    let explanation = config.explanation || "Transformation applied.";

    try {
        switch (config.action) {
            case 'remove_rows': {
                const initialRowCount = newDataObject.data.length;
                newDataObject = transformRemoveRows(newDataObject, config.conditions);
                const rowsRemoved = initialRowCount - newDataObject.data.length;
                return { success: true, message: `Removed ${rowsRemoved} row(s). ${explanation}`, newData: newDataObject };
            }
            case 'create_column': {
                if (newDataObject.meta.fields.includes(config.newColumn)) {
                    return { success: false, message: `Column "${config.newColumn}" already exists.`};
                }
                newDataObject = transformCreateColumn(newDataObject, config);
                // Add type for new column
                newDataObject.meta.inferredTypes[config.newColumn] = 'numerical';
                return { success: true, message: `Created new column '${config.newColumn}'. ${explanation}`, newData: newDataObject };
            }
            case 'rename_column': {
                if (!newDataObject.meta.fields.includes(config.oldColumn)) {
                    return { success: false, message: `Column to rename "${config.oldColumn}" does not exist.` };
                }
                if (newDataObject.meta.fields.includes(config.newColumn)) {
                    return { success: false, message: `A column named "${config.newColumn}" already exists.` };
                }
                newDataObject = transformRenameColumn(newDataObject, config);
                 // Update type mapping
                newDataObject.meta.inferredTypes[config.newColumn] = newDataObject.meta.inferredTypes[config.oldColumn];
                delete newDataObject.meta.inferredTypes[config.oldColumn];
                return { success: true, message: `Renamed column "${config.oldColumn}" to "${config.newColumn}". ${explanation}`, newData: newDataObject };
            }
            default:
                 return { success: false, message: `Unsupported transformation action: ${config.action}` };
        }
    } catch (error) {
        return { success: false, message: error.message };
    }
}

// --- Modular Transformation & Analysis Functions ---

function transformRemoveRows(dataObject, conditions) {
    const checkCondition = (row, condition) => {
        const { column, operator, value } = condition;
        const rowValue = row[column];
        if (rowValue === undefined) return false;
        const numericRowValue = parseFloat(String(rowValue).replace(/,/g, ''));
        const numericValue = typeof value === 'number' ? value : parseFloat(String(value).replace(/,/g, ''));
        switch (operator) {
            case 'equals': return rowValue == value;
            case 'not_equals': return rowValue != value;
            case 'greater_than': return !isNaN(numericRowValue) && !isNaN(numericValue) && numericRowValue > numericValue;
            case 'less_than': return !isNaN(numericRowValue) && !isNaN(numericValue) && numericRowValue < numericValue;
            case 'contains': return String(rowValue).toLowerCase().includes(String(value).toLowerCase());
            default: return false;
        }
    };
    dataObject.data = dataObject.data.filter(row => !conditions.every(cond => checkCondition(row, cond)));
    return dataObject;
}

function transformCreateColumn(dataObject, params) {
    const { newColumn, column1, operator, column2 } = params;
    dataObject.data = dataObject.data.map(row => {
        const val1 = parseFloat(String(row[column1]).replace(/,/g, ''));
        const val2 = parseFloat(String(row[column2]).replace(/,/g, ''));
        let newValue = null;
        if (!isNaN(val1) && !isNaN(val2)) {
            switch (operator) {
                case 'add': newValue = val1 + val2; break;
                case 'subtract': newValue = val1 - val2; break;
                case 'multiply': newValue = val1 * val2; break;
                case 'divide': newValue = val2 !== 0 ? (val1 / val2) : null; break;
            }
        }
        return { ...row, [newColumn]: newValue };
    });
    dataObject.meta.fields.push(newColumn);
    return dataObject;
}

function transformRenameColumn(dataObject, params) {
    const { oldColumn, newColumn } = params;
    dataObject.data = dataObject.data.map(row => {
        if (Object.prototype.hasOwnProperty.call(row, oldColumn)) {
            row[newColumn] = row[oldColumn];
            delete row[oldColumn];
        }
        return row;
    });
    const oldColumnIndex = dataObject.meta.fields.indexOf(oldColumn);
    if (oldColumnIndex > -1) {
        dataObject.meta.fields[oldColumnIndex] = newColumn;
    }
    return dataObject;
}

function executeAggregate(dataObject, params) {
    const { groupBy, aggregation, valueColumn } = params;
    const groups = {};

    dataObject.data.forEach(row => {
        const key = row[groupBy];
        if (key === undefined || key === null) return;
        
        groups[key] = groups[key] || { sum: 0, count: 0 };
        
        if (aggregation !== 'count') {
            const value = parseFloat(String(row[valueColumn]).replace(/,/g, ''));
            if (!isNaN(value)) {
                groups[key].sum += value;
            }
        }
        groups[key].count += 1;
    });

    const aggregatedData = [];
    const newColumnName = `${aggregation}_${valueColumn || 'count'}`;

    for (const key in groups) {
        let value;
        if (aggregation === 'sum') {
            value = groups[key].sum;
        } else if (aggregation === 'average') {
            value = groups[key].count > 0 ? groups[key].sum / groups[key].count : 0;
        } else { // count
            value = groups[key].count;
        }
        aggregatedData.push({ [groupBy]: key, [newColumnName]: value });
    }

    return {
        data: aggregatedData,
        meta: { fields: [groupBy, newColumnName] }
    };
}

function executeSort(dataObject, params) {
    const { column, order } = params;
    dataObject.data.sort((a, b) => {
        const valA = a[column];
        const valB = b[column];
        const numA = parseFloat(valA);
        const numB = parseFloat(valB);

        let comparison = 0;
        if (!isNaN(numA) && !isNaN(numB)) {
            comparison = numA - numB;
        } else {
            comparison = String(valA).localeCompare(String(valB));
        }
        return order === 'descending' ? -comparison : comparison;
    });
    return dataObject;
}

function executeLimit(dataObject, params) {
    dataObject.data = dataObject.data.slice(0, params.count);
    return dataObject;
}


async function processAiResponse(responseText) {
    const parts = responseText.split('---VIZ---');
    const textAnswer = parts[0].trim();
    if (textAnswer) {
        addMessageToChat(textAnswer, 'ai');
    }

    if (parts.length > 1 && parts[1].trim() !== 'NONE') {
        try {
            const vizJsonString = parts[1].trim().replace(/```json/g, '').replace(/```/g, '');
            const vizResponse = JSON.parse(vizJsonString);

            if (vizResponse.visualizations && Array.isArray(vizResponse.visualizations)) {
                // Handle vague queries that return multiple charts
                for (const viz of vizResponse.visualizations) {
                    if (viz.analysisType === 'complex') {
                        await addChartMessageToChat(viz.chartConfig, viz.chartData);
                    } else { // 'simple' or old format without analysisType
                        await addChartMessageToChat(viz.chartConfig || viz, null);
                    }
                }
            } else {
                // Handle single chart response
                if (vizResponse.analysisType === 'complex') {
                    await addChartMessageToChat(vizResponse.chartConfig, vizResponse.chartData);
                } else { // 'simple' or old format
                    const chartConfig = vizResponse.chartConfig || vizResponse;
                    await addChartMessageToChat(chartConfig, null);
                }
            }
        } catch (error) {
            addErrorMessageToChat('Visualization Error', `I couldn't create a chart from the data. The AI's response may have been malformed. Details: ${error.message}`);
            console.error("Visualization JSON parsing error:", error, parts[1].trim());
        }
    }
    await generateAndDisplayFollowUpQuestions('<strong>What\'s next?</strong> Here are some ideas:');
}


function addMessageToChat(message, sender) {
    const messageEl = document.createElement('div');
    messageEl.classList.add('message', `${sender}-message`);
    messageEl.innerHTML = message;
    
    if (isFilteredState && sender === 'ai') {
        messageEl.dataset.filterMessage = 'true';
    }

    chatHistory.appendChild(messageEl);
    chatHistory.scrollTop = chatHistory.scrollHeight;
}

function addErrorMessageToChat(title, message) {
    const messageEl = document.createElement('div');
    messageEl.classList.add('message', 'error-message');
    messageEl.innerHTML = `<strong>${title}</strong><p>${message}</p>`;
    chatHistory.appendChild(messageEl);
    chatHistory.scrollTop = chatHistory.scrollHeight;
    removeLoadingIndicator();
}

async function getColumnNameSuggestion(invalidColumn, availableColumns) {
    if (!ai || !invalidColumn || !availableColumns.length) {
        return null;
    }

    const prompt = `
You are a helpful assistant. A user's query for a data visualization failed because they used a column name that doesn't exist. Your task is to find the most likely intended column from the list of available columns.

User's column name: "${invalidColumn}"
Available columns in the dataset: [${availableColumns.join(', ')}]

Which of the available columns is the best match for the user's column name?

Your response MUST be one of two things:
1. The single, best-matching column name from the "Available columns" list.
2. The word "NONE" if no column is a reasonably close match (e.g., the names are completely different).

Do not add any explanation or other text.`;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
        });
        const suggestion = response.text.trim();
        // Check if suggestion is valid before returning
        if (suggestion.toUpperCase() !== 'NONE' && availableColumns.find(c => c.trim().toLowerCase() === suggestion.trim().toLowerCase())) {
            // Find the original casing of the column to be precise
            return availableColumns.find(c => c.trim().toLowerCase() === suggestion.trim().toLowerCase());
        }
        return null;
    } catch (error) {
        console.error("Error getting column name suggestion:", error);
        return null;
    }
}


async function addChartMessageToChat(config, precomputedData = null) {
    const chartId = `chart-${Date.now()}-${Math.random()}`;
    const messageEl = document.createElement('div');
    messageEl.classList.add('message', 'ai-message', 'chart-message');

    if (isFilteredState) {
        messageEl.dataset.filterMessage = 'true';
    }
    
    const titleEl = document.createElement('p');
    titleEl.innerHTML = `<strong>${config.title || 'Visualization'}</strong>`;
    
    const chartWrapper = document.createElement('div');
    chartWrapper.classList.add('chart-wrapper');

    const buttonContainer = document.createElement('div');
    buttonContainer.classList.add('chart-actions');

    const exportButton = document.createElement('button');
    exportButton.textContent = 'Export PNG';
    exportButton.classList.add('chart-action-btn');
    exportButton.onclick = () => exportChart(chartId, config.title);
    buttonContainer.appendChild(exportButton);

    const isZoomable = ['bar', 'line', 'area', 'scatter'].includes(config.chartType);
    if (isZoomable) {
        const resetZoomButton = document.createElement('button');
        resetZoomButton.textContent = 'Reset Zoom';
        resetZoomButton.classList.add('chart-action-btn');
        resetZoomButton.onclick = () => {
            const chartInstance = chartInstances.get(chartId);
            if (chartInstance) {
                chartInstance.resetZoom();
            }
        };
        buttonContainer.appendChild(resetZoomButton);
    }

    const canvas = document.createElement('canvas');
    canvas.id = chartId;
    
    chartWrapper.appendChild(buttonContainer);
    chartWrapper.appendChild(canvas);

    messageEl.appendChild(titleEl);
    messageEl.appendChild(chartWrapper);

    if (config.explanation) {
        const explanationEl = document.createElement('div');
        explanationEl.classList.add('chart-explanation');
        explanationEl.innerHTML = config.explanation;
        messageEl.appendChild(explanationEl);
    }
    
    chatHistory.appendChild(messageEl);
    chatHistory.scrollTop = chatHistory.scrollHeight;

    try {
        renderVisualization(config, canvas, precomputedData);
    } catch (error) {
        console.error("Failed to render chart:", error);
        // Remove the chart container and explanation if rendering fails
        messageEl.remove();

        if (error.message.startsWith('Invalid column name')) {
            const match = error.message.match(/Could not find a match for "([^"]+)"/);
            const invalidColumn = match ? match[1] : null;

            if (invalidColumn) {
                const suggestion = await getColumnNameSuggestion(invalidColumn, activeData.meta.fields);
                if (suggestion) {
                    const suggestionMsg = `It looks like the column "<em>${invalidColumn}</em>" doesn't exist. Did you mean "<strong>${suggestion}</strong>"?<br><br>Please try your query again with the correct column name.`;
                    addErrorMessageToChat('Column Not Found', suggestionMsg);
                } else {
                    addErrorMessageToChat('Chart Rendering Failed', error.message);
                }
            } else {
                addErrorMessageToChat('Chart Rendering Failed', error.message);
            }
        } else {
            addErrorMessageToChat('Chart Rendering Failed', error.message);
        }
    }
}


function showLoadingIndicator() {
    const loadingEl = document.createElement('div');
    loadingEl.classList.add('message', 'ai-message', 'loading-indicator');
    loadingEl.id = 'loading';
    loadingEl.innerHTML = '<span></span><span></span><span></span>';
    chatHistory.appendChild(loadingEl);
    chatHistory.scrollTop = chatHistory.scrollHeight;
}

function removeLoadingIndicator() {
    const loadingEl = document.getElementById('loading');
    if (loadingEl) {
        loadingEl.remove();
    }
}

// --- UI Helpers ---

/**
 * Automatically adjusts the height of the chat textarea based on its content.
 */
function autoResizeTextarea(e) {
    const textarea = e.target;
    textarea.style.height = 'auto'; // Reset height to recalculate scrollHeight
    textarea.style.height = `${textarea.scrollHeight}px`;
}

/**
 * Sets up a ResizeObserver to watch the chat input container's height
 * and updates a CSS custom property to adjust the layout accordingly,
 * preventing content from being hidden behind the input box.
 */
function observeChatInputResize() {
    const chatInputContainer = document.getElementById('chat-input-container');
    if (!chatInputContainer) return;

    const resizeObserver = new ResizeObserver(entries => {
        for (let entry of entries) {
            const height = entry.contentRect.height;
            // Set the CSS custom property used for dynamic padding/margins
            document.documentElement.style.setProperty('--chat-input-height', `${height}px`);
        }
    });

    resizeObserver.observe(chatInputContainer);
}


// --- Visualization ---

/**
 * Creates and shows a temporary tooltip on the page.
 * @param {number} x The horizontal coordinate.
 * @param {number} y The vertical coordinate.
 * @param {string} text The text to display in the tooltip.
 */
function showClickTooltip(x, y, text) {
    hideClickTooltip(); // Ensure no other tooltips are visible

    const tooltip = document.createElement('div');
    tooltip.className = 'click-tooltip';
    tooltip.textContent = text;
    document.body.appendChild(tooltip);
    currentTooltip = tooltip;

    tooltip.style.left = `${x}px`;
    tooltip.style.top = `${y}px`;

    requestAnimationFrame(() => {
        tooltip.style.opacity = '1';
    });

    setTimeout(hideClickTooltip, 2500); // Automatically hide after 2.5 seconds
}

/**
 * Fades out and removes the currently visible tooltip.
 */
function hideClickTooltip() {
    if (currentTooltip) {
        currentTooltip.style.opacity = '0';
        setTimeout(() => {
            if (currentTooltip) {
                currentTooltip.remove();
                currentTooltip = null;
            }
        }, 200); // Delay removal to allow for fade-out transition
    }
}


/**
 * Finds a matching column name from the actual data columns, ignoring case and whitespace.
 * @param {string} aiColumnName The column name provided by the AI.
 * @param {string[]} actualColumns The array of actual column names from the CSV header.
 * @returns {string|null} The matching actual column name, or null if not found.
 */
function findMatchingColumn(aiColumnName, actualColumns) {
    if (!aiColumnName) return null;
    const normalizedAiName = aiColumnName.trim().toLowerCase();
    
    const foundColumn = actualColumns.find(col => col.trim().toLowerCase() === normalizedAiName);
    return foundColumn || null;
}

function renderVisualization(config, canvasElement, precomputedData = null) {
    let chartData;
    let axisTitles;

    // Neobrutalism Palette
    const colors = [
        '#4285F4', // Blue
        '#EA4335', // Red
        '#FBBC05', // Yellow
        '#34A853', // Green
        '#A142F4', // Purple
        '#24C1E0'  // Cyan
    ];
    const borderColor = '#000000';
    const borderWidth = 2;
    const textColor = '#000000';
    const gridColor = '#e0e0e0';

    if (precomputedData) {
        axisTitles = config.axisTitles || { x: null, y: null };
        const dataset = {
            label: axisTitles.y || config.title || '',
            data: precomputedData.values,
            borderWidth: borderWidth,
            borderColor: borderColor,
        };
        
        if (config.chartType === 'line' || config.chartType === 'area') {
            dataset.backgroundColor = 'rgba(66, 133, 244, 0.2)'; // Light blue fill
            dataset.borderColor = colors[0];
            dataset.pointBackgroundColor = colors[0];
            dataset.pointBorderColor = '#000';
            dataset.pointRadius = 4;
            dataset.pointHoverRadius = 6;
            dataset.tension = 0.2;
            if (config.chartType === 'area') {
                dataset.fill = true; 
            }
        } else {
            dataset.backgroundColor = colors;
        }

        chartData = {
            labels: precomputedData.labels,
            datasets: [dataset]
        };

    } else {
        if (!activeData) return;
    
        const prepared = prepareChartData(config);
        if (!prepared) {
            throw new Error('Could not prepare data for the requested chart type.');
        }
        chartData = prepared.data;
        axisTitles = prepared.titles;
    }


    if (chartInstances.has(canvasElement.id)) {
        chartInstances.get(canvasElement.id).destroy();
    }

    const chartTypeForChartJS = config.chartType === 'area' ? 'line' : (config.chartType === 'donut' ? 'doughnut' : config.chartType);
    
    const options = {
        responsive: true,
        maintainAspectRatio: false,
        onClick: (e, elements, chart) => {
            if (elements.length === 0) return;

            if (precomputedData) {
                showClickTooltip(e.native.clientX, e.native.clientY, "Filtering is not available for this custom chart.");
                return;
            }

            if (clickTimer) {
                clearTimeout(clickTimer);
                clickTimer = null;
                hideClickTooltip();

                const elementIndex = elements[0].index;
                const clickedLabel = chart.data.labels[elementIndex];
                
                const labelColName = ['pie', 'donut', 'doughnut'].includes(chart.config.type)
                    ? config.categoryColumn
                    : config.xAxisColumn;
                
                const filterColumn = findMatchingColumn(labelColName, activeData.meta.fields);

                if (filterColumn && clickedLabel) {
                    const newFilter = { column: filterColumn, value: clickedLabel };
                    const isDuplicate = stagedFilters.some(f => f.column === newFilter.column && f.value === newFilter.value);
                    const isAlreadyActive = activeFilters.some(f => f.column === newFilter.column && f.value === newFilter.value);

                    if (!isDuplicate && !isAlreadyActive) {
                        stagedFilters.push(newFilter);
                        updateStagingUi();
                    } else {
                        showClickTooltip(e.native.clientX, e.native.clientY, "This filter is already active or staged.");
                    }
                }
            } else {
                clickTimer = setTimeout(() => {
                    clickTimer = null;
                    const value = chart.data.labels[elements[0].index];
                    showClickTooltip(e.native.clientX, e.native.clientY, `Double-click to stage filter for "${value}"`);
                }, CLICK_DELAY);
            }
        },
        plugins: {
            title: {
                display: false,
            },
            legend: {
                labels: { 
                    color: textColor,
                    font: { family: "'JetBrains Mono', monospace", size: 12 },
                    boxWidth: 15,
                    usePointStyle: true,
                    pointStyle: 'rectRounded'
                }
            },
            tooltip: {
                enabled: true,
                backgroundColor: '#fff',
                titleColor: '#000',
                bodyColor: '#000',
                borderColor: '#000',
                borderWidth: 2,
                titleFont: { size: 14, weight: 'bold', family: "'Inter', sans-serif" },
                bodyFont: { size: 12, family: "'JetBrains Mono', monospace" },
                padding: 10,
                cornerRadius: 4,
                displayColors: true,
                boxPadding: 4,
            }
        },
        scales: (config.chartType !== 'pie' && config.chartType !== 'doughnut' && config.chartType !== 'donut') ? {
            x: { 
                ticks: { color: textColor, font: { family: "'JetBrains Mono', monospace" } }, 
                grid: { color: gridColor, drawBorder: false },
                border: { display: true, color: '#000', width: 2 },
                title: {
                    display: !!axisTitles.x,
                    text: axisTitles.x || '',
                    color: textColor,
                    font: { size: 14, weight: '800', family: "'Inter', sans-serif" }
                }
            },
            y: { 
                ticks: { color: textColor, font: { family: "'JetBrains Mono', monospace" } }, 
                grid: { color: gridColor, drawBorder: false },
                border: { display: true, color: '#000', width: 2 },
                title: {
                    display: !!axisTitles.y,
                    text: axisTitles.y || '',
                    color: textColor,
                    font: { size: 14, weight: '800', family: "'Inter', sans-serif" }
                }
            }
        } : {}
    };

    const isZoomable = ['bar', 'line', 'area', 'scatter'].includes(config.chartType);
    if (isZoomable) {
        options.plugins.zoom = {
            pan: {
                enabled: true,
                mode: 'xy',
                threshold: 5,
            },
            zoom: {
                wheel: { enabled: true },
                pinch: { enabled: true },
                mode: 'xy',
            }
        };
    }

    const ctx = canvasElement.getContext('2d');
    const newChartInstance = new Chart(ctx, {
        type: chartTypeForChartJS,
        data: chartData,
        options: options,
    });
    
    chartInstances.set(canvasElement.id, newChartInstance);
}

function prepareChartData(config) {
    const { data, meta } = activeData;
    const { chartType, xAxisColumn, yAxisColumn, categoryColumn, valueColumn, aggregation } = config;
    const actualColumns = meta.fields;

    let finalChartData;
    const axisTitles = { x: null, y: null };
    
    // Neobrutalism Chart Colors
    const colors = [
        '#4285F4', // Blue
        '#EA4335', // Red
        '#FBBC05', // Yellow
        '#34A853', // Green
        '#A142F4', // Purple
        '#24C1E0'  // Cyan
    ];
    
    if (['bar', 'line', 'area', 'pie', 'donut'].includes(chartType)) {
        const aiLabelCol = ['pie', 'donut'].includes(chartType) ? categoryColumn : xAxisColumn;
        const aiValueCol = ['pie', 'donut'].includes(chartType) ? valueColumn : yAxisColumn;

        const labelCol = findMatchingColumn(aiLabelCol, actualColumns);
        
        // Value column is optional for 'count' aggregation
        const valueCol = aggregation === 'count' ? null : findMatchingColumn(aiValueCol, actualColumns);

        if (!labelCol) {
            let errorMsg = `Invalid column name for chart labels. Could not find a match for "${aiLabelCol}". Available columns are: [${actualColumns.join(', ')}]`;
            throw new Error(errorMsg);
        }
        if (aggregation !== 'count' && !valueCol) {
            let errorMsg = `Invalid column name for chart values. Could not find a match for "${aiValueCol}". Available columns are: [${actualColumns.join(', ')}]`;
            throw new Error(errorMsg);
        }

        let aggregated;

        // aggregation can be 'average', 'sum', or 'count'
        if (aggregation === 'average' && valueCol) {
            axisTitles.x = labelCol;
            axisTitles.y = `Average of ${valueCol}`;
            const sums = {};
            const counts = {};
            data.forEach(row => {
                const key = row[labelCol];
                const value = parseFloat(String(row[valueCol]).replace(/,/g, ''));
                if (key && !isNaN(value)) {
                    sums[key] = (sums[key] || 0) + value;
                    counts[key] = (counts[key] || 0) + 1;
                }
            });
            aggregated = {};
            for (const key in sums) {
                if (counts[key] > 0) {
                    aggregated[key] = sums[key] / counts[key];
                }
            }
        } else if (aggregation === 'sum' && valueCol) {
            axisTitles.x = labelCol;
            axisTitles.y = `Total ${valueCol}`;
            aggregated = data.reduce((acc, row) => {
                const key = row[labelCol];
                const value = parseFloat(String(row[valueCol]).replace(/,/g, ''));
                if (key && !isNaN(value)) {
                    acc[key] = (acc[key] || 0) + value;
                }
                return acc;
            }, {});
        } else { // This handles aggregation: 'count' where valueCol is omitted
            axisTitles.x = labelCol;
            axisTitles.y = 'Count';
            aggregated = data.reduce((acc, row) => {
                const key = row[labelCol];
                if (key) {
                    acc[key] = (acc[key] || 0) + 1;
                }
                return acc;
            }, {});
        }
        
        // Convert aggregated data to an array for intelligent sorting.
        let allItems = Object.entries(aggregated);

        // Apply common-sense sorting for graceful presentation.
        if (['bar', 'pie', 'donut'].includes(chartType)) {
            // For categorical charts, sort by value (descending) to show the most impactful items first.
            allItems.sort(([, a], [, b]) => b - a);
        } else if (['line', 'area'].includes(chartType)) {
            // For sequential charts (like time series), sort by the label to maintain a logical order.
            allItems.sort(([a], [b]) => {
                const dateA = new Date(a);
                const dateB = new Date(b);

                // If labels can be parsed as valid dates, sort chronologically.
                if (!isNaN(dateA.getTime()) && !isNaN(dateB.getTime())) {
                    return dateA - dateB;
                }
                
                // Otherwise, use natural sorting for strings containing numbers (e.g., "Q1", "Q10").
                return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
            });
        }
        
        let labels, values;

        // Automatically group numerous categories into 'Other' for clarity, especially after sorting.
        const MAX_CATEGORIES_TO_DISPLAY = 20;
        if (['bar', 'pie', 'donut'].includes(chartType) && allItems.length > MAX_CATEGORIES_TO_DISPLAY) {
            const topItems = allItems.slice(0, MAX_CATEGORIES_TO_DISPLAY - 1);
            const otherItems = allItems.slice(MAX_CATEGORIES_TO_DISPLAY - 1);
            
            const otherSum = otherItems.reduce((sum, [, value]) => sum + value, 0);
            
            labels = topItems.map(([key]) => key);
            values = topItems.map(([, value]) => value);

            labels.push('Other');
            values.push(otherSum);
        } else {
            labels = allItems.map(([key]) => key);
            values = allItems.map(([, value]) => value);
        }


        const dataset = {
            label: valueCol ? (aggregation === 'average' ? `Average of ${valueCol}`: `Total ${valueCol}`) : 'Count',
            data: values,
            borderWidth: 2,
            borderColor: '#000', // Black border
        };

        if (chartType === 'line' || chartType === 'area') {
            dataset.backgroundColor = 'rgba(66, 133, 244, 0.2)';
            dataset.borderColor = colors[0];
            dataset.pointBackgroundColor = colors[0];
            dataset.pointBorderColor = '#000';
            dataset.pointRadius = 4;
            dataset.tension = 0.2;
            if (chartType === 'area') {
                dataset.fill = true; 
            }
        } else {
            dataset.backgroundColor = colors;
        }

        finalChartData = {
            labels: labels,
            datasets: [dataset]
        };
    } 
    else if (chartType === 'scatter') {
        const matchedXAxisColumn = findMatchingColumn(xAxisColumn, actualColumns);
        const matchedYAxisColumn = findMatchingColumn(yAxisColumn, actualColumns);

        if (!matchedXAxisColumn || !matchedYAxisColumn) {
            let errorMsg = `Invalid column names for scatter plot. `;
            if (!matchedXAxisColumn) errorMsg += `Could not find a match for column "${xAxisColumn}". `;
            if (!matchedYAxisColumn) errorMsg += `Could not find a match for column "${yAxisColumn}". `;
            errorMsg += `Available columns are: [${actualColumns.join(', ')}]`;
            throw new Error(errorMsg);
        }
        
        axisTitles.x = matchedXAxisColumn;
        axisTitles.y = matchedYAxisColumn;

        const points = data.map(row => ({
            x: parseFloat(String(row[matchedXAxisColumn]).replace(/,/g, '')),
            y: parseFloat(String(row[matchedYAxisColumn]).replace(/,/g, ''))
        })).filter(p => !isNaN(p.x) && !isNaN(p.y));
        
        finalChartData = {
            datasets: [{
                label: `${matchedYAxisColumn} vs ${matchedXAxisColumn}`,
                data: points,
                backgroundColor: colors[0],
                borderColor: '#000',
                borderWidth: 1,
                pointRadius: 6,
                pointHoverRadius: 8
            }]
        };
    } else {
        return null;
    }
    
    return { data: finalChartData, titles: axisTitles };
}


function exportChart(chartId, title) {
    const chartInstance = chartInstances.get(chartId);
    if (chartInstance) {
        const link = document.createElement('a');
        link.href = chartInstance.toBase64Image();
        const sanitizedTitle = title ? title.replace(/[^a-z0-9]/gi, '_').toLowerCase() : 'chart';
        link.download = `${sanitizedTitle}.png`;
        link.click();
    }
}

// --- Data Export ---
function exportAsCsv() {
    if (!activeData || !activeData.data) {
        addErrorMessageToChat('Export Failed', 'No data is available to export.');
        return;
    }
    const csv = Papa.unparse(activeData.data);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    const fileName = (fileNameEl.textContent || 'data_export').replace(/\.(csv|xlsx|xls)$/, '');
    link.setAttribute("download", `${fileName}_transformed.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function exportAsExcel() {
    if (!activeData || !activeData.data) {
        addErrorMessageToChat('Export Failed', 'No data is available to export.');
        return;
    }
    const worksheet = XLSX.utils.json_to_sheet(activeData.data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
    const fileName = (fileNameEl.textContent || 'data_export').replace(/\.(csv|xlsx|xls)$/, '');
    XLSX.writeFile(workbook, `${fileName}_transformed.xlsx`);
}