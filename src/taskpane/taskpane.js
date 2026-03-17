import { Remarkable } from 'remarkable';
console.log("taskpane.js loaded");

// INIT GlOBAL EXCEL CONTEXT ---
let chatContext = [];
let cellContext = "";

// Office.js initialization with better error handling
function initializeOfficeApp() {
    // Check if Office.js is loaded
    if (typeof Office !== 'undefined') {
        console.log("Office.js detected, calling onReady");
        Office.onReady((info) => {
            console.log("Office.onReady fired, host:", info.host);
            if (info.host === Office.HostType.Excel) {
                console.log("Setting up event handlers for Excel");
                // Ensure DOM is ready before accessing elements
                if (document.readyState === 'loading') {
                    document.addEventListener('DOMContentLoaded', setupEventHandlers);
                } else {
                    setupEventHandlers();
                }
            } else {
                console.log("Not Excel host:", info.host);
            }
        }).catch(error => {
            console.error("Error in Office.onReady:", error);
            // Retry after a short delay if there's an error
            setTimeout(initializeOfficeApp, 1000);
        });
    } else {
        console.error("Office object is not defined. Please ensure Office.js is loaded.");
        // Retry after a short delay
        setTimeout(initializeOfficeApp, 1000);
    }
}

function setupEventHandlers(retryCount = 0) {
    try {
        // Wait for all required elements to be available
        const sendBtn = document.getElementById("send-btn");
        const newChatBtn = document.getElementById("new-chat");
        const chatHistory = document.getElementById("chat-history");
        const sideloadMsg = document.getElementById("sideload-msg");
        const appBody = document.getElementById("app-body");
        
        // If elements are not yet available, retry (max 20 retries)
        if (!sendBtn || !newChatBtn || !chatHistory || !sideloadMsg || !appBody) {
            if (retryCount < 20) {
                console.log("Some DOM elements not found, retrying... (attempt " + (retryCount + 1) + ")");
                setTimeout(() => setupEventHandlers(retryCount + 1), 100);
                return;
            } else {
                console.error("Failed to find required DOM elements after 20 attempts");
                return;
            }
        }
        
        // Set up event handlers
        sendBtn.onclick = sendToAI;
        newChatBtn.onclick = () => {
            chatContext = [];
            if (chatHistory) {
                chatHistory.innerHTML = "";
            }
        };
        
        console.log("Toggling app body visibility");
        console.log("sideload-msg element found:", !!sideloadMsg);
        console.log("app-body div element found:", !!appBody);
        if (sideloadMsg) {
            sideloadMsg.style.display = "none";
        }
        if (appBody) {
            appBody.style.display = "flex";
        }
        console.log("Visibility toggle complete");
    } catch (error) {
        console.error("Error setting up event handlers:", error);
        // Retry after a short delay
        setTimeout(setupEventHandlers, 500);
    }
}

// Initialize the app when the page loads
function initializeApp() {
    // Check if Office.js is already loaded
    if (typeof Office !== 'undefined') {
        initializeOfficeApp();
    } else {
        // If Office.js is not loaded yet, wait for it
        const checkOfficeInterval = setInterval(() => {
            if (typeof Office !== 'undefined') {
                clearInterval(checkOfficeInterval);
                initializeOfficeApp();
            }
        }, 100);
        
        // Timeout after 10 seconds to prevent infinite waiting
        setTimeout(() => {
            clearInterval(checkOfficeInterval);
            console.error("Office.js failed to load within 10 seconds");
            // Even if Office.js fails to load, try to set up event handlers anyway
            // This can help in some edge cases
            if (document.readyState === 'loading') {
                document.addEventListener('DOMContentLoaded', setupEventHandlers);
            } else {
                setupEventHandlers();
            }
        }, 10000);
    }
}

// Initialize the app when the page loads
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeApp);
} else {
    initializeApp();
}

async function sendToAI() {
    const promptInput = document.getElementById("user-prompt");
    const sendBtn = document.getElementById("send-btn");
    const prompt = promptInput.value;
    const model = document.getElementById("model-select").value;
    const chatHistory = document.getElementById("chat-history");
    const apiKey = process.env.OPENROUTER_API_KEY;

    if (!prompt) {
        appendChat("Error", "Please enter a prompt before sending.");
        return;
    }

    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load("address");
            range.load("values"); // Load the data from the cells
            await context.sync();
            
            const cellReference = range.address;
            // Convert 2D array of values into a readable string
            const values = range.values;
            const cellValue = values.map(row => row.join("\t")).join("\n");
            cellContext = `"cellReference":"${cellReference}","cellValue":"${cellValue}"`;
        });
    } catch (error) {
        console.warn("Could not read Excel selection:", error);
    }

    // Check if API key is available
    if (!apiKey || apiKey === 'your_key_here') {
        appendChat("Error", "API key not configured. Please set your OPENROUTER_API_KEY in the .env file.");
        return;
    }

    // Disable button and show loading state
    const originalButtonText = sendBtn.textContent;
    sendBtn.disabled = true;
    sendBtn.textContent = "Sending...";

    appendChat("User", prompt);
    promptInput.value = "";

    // Add "Thinking..." indicator
    const thinkingMsg = document.createElement("div");
    thinkingMsg.className = "ai thinking-mode";
    thinkingMsg.innerHTML = "<em>AI is thinking...</em>";
    chatHistory.appendChild(thinkingMsg);
    chatHistory.scrollTop = chatHistory.scrollHeight;
    thinkingFlag = true;

    const systemInstructions = `
    You are an expert Microsoft Excel assistant.
    - If user does not specify the solution approach, arrange solutions approach priority as below:
    1. Standard Excel UI step-by-step but concise operations (Menus/Ribbon and corresponding Keyboard shortcuts).
    2. Excel Functions/Formulas.
    3. VBA Scripts (only consider if the above cannot solve it).
    4. Power Query M (only consider if the above cannot solve it).

    - Response Constraints
    1. Give response directly, no repeat on prompts.
    2. Response involved 1-2 solution approaches, no more than 2.
    3. Provide suggestion at the end of the response, if applicable.
    4. No conclusion at the end of the response.
    5. If the operations or Functions/Formulas approaches are provided, do NOT provide VBA Scripts solution approach.
    6. If the operations or Functions/Formulas or VBA Scripts approaches are provided, do NOT provide Power Query M solution approach.
    
    - Additional Guidelines
    1. For Excel Functions/Formulas or VBA scripts, explain how it works in concise bullet points.
    2. Add Notes & variations in case it needs.
    3. Test the functions/formulas/scripts first to ensure they're executable before providing.
    4. Include office/excel version required for the solution.
    
    - Format of Responses
    1. Use markdown format for responses.
    2. Markdown start from heading level 3 ###.
    3. If providing an Excel formula, wrap it in \`\`\`excel blocks.
    4. If providing VBA code, wrap it in \`\`\`vba blocks.
    5. If providing Power Query M, wrap it in \`\`\`M blocks.
    6. For step-by-step operations, provide them in a numbered list.
    `;
    
    // Combine context and prompt
    const promptWithCellContext = cellContext ? `{${cellContext}, "question":"${prompt}"}` : `{"question":"${prompt}"}`
    const contextString = chatContext.length > 0 ? JSON.stringify(chatContext) : "";
    const userPrompt = contextString 
        ? `Context:\n{${contextString}}\n\nUserQuestion: ${promptWithCellContext}`
        : `${promptWithCellContext}`;
    console.log("--- DEBUG: Context (https://playcode.io/json-formatter) ---");
    console.log(contextString || "[Empty Selection]");
    console.log("--- DEBUG: UserQuestion ---");
    console.log(promptWithCellContext || "[Empty Selection]");

    try {
        // CORRECTED URL
        const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${apiKey}`,
                "Content-Type": "application/json",
                "HTTP-Referer": "http://localhost:3000", // Required by some OpenRouter models
                "X-Title": "ExcelAI"
            },
            body: JSON.stringify({
                model: model,
                stream: true,  // Enable streaming responses
                messages: [
                    { role: "system", content: systemInstructions },
                    { role: "user", content: userPrompt }
                ]
            })
        });

        // Check if response is ok
        if (!response.ok) {
            // NOTE: Only call .json() here because the stream hasn't been started yet
            const errorData = await response.json();
            throw new Error(errorData.error?.message || `HTTP ${response.status}`);
        }

        // PROCESS THE READABLE STREAM
        const reader = response.body.getReader();
        const decoder = new TextDecoder("utf-8");
        let fullText = "";

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            const chunk = decoder.decode(value);
            const lines = chunk.split("\n");
            
            for (const line of lines) {
                if (line.startsWith("data: ")) {
                    const data = line.slice(6);
                    if (data === "[DONE]") break;
                    try {
                        const json = JSON.parse(data);
                        const content = json.choices[0].delta.content || "";
                        if (content) { // CRITICAL: Only update if content exists
                            if (thinkingFlag === true) {
                                // Remove thinking indicator as data starts arriving
                                if (chatHistory.contains(thinkingMsg)) chatHistory.removeChild(thinkingMsg);
                                thinkingFlag = false;
                            }
                            fullText += content;
                            // CALL 1: Streaming update (fast, no buttons yet)
                            appendChat("AI", fullText, false); 
                        }
                    } catch (e) { /* Ignore partial JSON chunks */ }
                }
            }
        }
        // CALL 2: Final cleanup (processes Markdown and adds buttons)
        appendChat("AI", fullText, true);
        
        // memory for further chat
        chatContext.push(`{"historyQuestion":"${promptWithCellContext}","historyResponse":"${fullText}"}`);

    } catch (error) {
        if (chatHistory.contains(thinkingMsg)) chatHistory.removeChild(thinkingMsg);
        console.error("Error in sendToAI:", error);
        appendChat("Error", error.message);
    } finally {
        // Re-enable button and restore text
        sendBtn.disabled = false;
        sendBtn.textContent = originalButtonText;
    }
}

function codeToBase64(str) {
    try { 
        // 1. Convert string to UTF-8 bytes
        const bytes = new TextEncoder().encode(str);
        // 2. Convert bytes to a binary string
        const binString = Array.from(bytes, (byte) => String.fromCharCode(byte)).join("");
        // 3. Encode to Base64
        const encodeCode = btoa(binString);
        return encodeCode;
    } catch (error) {
        console.error("Encode Error:", error);
    }
    return "";
}

function codeFromBase64(base64Code) {
    try {
        // 1. Convert Base64 back to a binary string
        const binString = atob(base64Code);
        // 2. Convert the binary string into a byte array (Uint8Array)
        const bytes = Uint8Array.from(binString, (char) => char.charCodeAt(0));
        // 3. Decode the byte array back into a UTF-8 string
        const decodedCode = new TextDecoder().decode(bytes);
        return decodedCode;
    } catch (error) {
        console.error("Decode Error:", error);
    }
    return "";
}

/**
 * Copies text content to the system clipboard
 * @param {string} text - The code to copy
 */
async function copyToClipboard(base64Code) {
    try {
        const text = codeFromBase64(base64Code);
        // Use the modern Clipboard API
        await navigator.clipboard.writeText(text);
        
        // Optional: Provide UI feedback
        const activeBtn = document.activeElement;
        if (activeBtn && activeBtn.tagName === "BUTTON") {
            const originalText = activeBtn.textContent;
            activeBtn.textContent = "Copied!";
            activeBtn.style.backgroundColor = "#217346";
            setTimeout(() => {
                activeBtn.textContent = originalText;
                activeBtn.style.backgroundColor = "#666";
            }, 2000);
        }
    } catch (err) {
        console.error("Failed to copy: ", err);
        appendChat("Error", "Failed to copy to clipboard.");
    }
}
// Make it globally accessible
window.copyToClipboard = copyToClipboard;

async function applyToExcel(base64Code) {
    try {
        await Excel.run(async (context) => {
            const code = codeFromBase64(base64Code);
            const range = context.workbook.getSelectedRange();
            
            // If the code starts with '=', treat it as a formula, otherwise as text
            if (code.startsWith("=")) {
                range.formulas = [[code]];
            } else {
                range.values = [[code]];
            }
            
            await context.sync();
        });
    } catch (error) {
        console.error("Error applying to Excel: ", error);
        // Optional: show error in the chat
        appendChat("Error", "Could not apply to cell: " + error.message);
    }
}
window.applyToExcel = applyToExcel;

// Helper function to generate code block HTML with action buttons
function createCodeBlockHTML(code, lang) {
    const language = lang || (code.startsWith('=') ? 'excel' : 'others');
    const base64Code = codeToBase64(code);
    
    let actionButton = "";
    if (language === "others") {
        actionButton = `<button class="apply-btn copy-btn" style="background-color: #666;" data-code="${base64Code}">Copy</button>`;
    } else {
        // Use data attribute instead of inline onclick
        actionButton = `<button class="apply-btn excel-btn" data-code="${base64Code}">Add to Excel</button>`;
    }
    
    return `
        <div class="code-block-container">
            <div class="code-header" style="font-size:0.7em; color:#aaa; margin-bottom:4px; text-transform:uppercase; padding: 8px 10px;">
                ${language}
            </div>
            <pre><code>${code}</code></pre>
            <div class="code-actions" style="padding: 8px 10px;">
                ${actionButton}
            </div>
        </div>`;
}
document.addEventListener('click', function(e) {
    if (e.target.classList.contains('copy-btn')) {
        const base64Code = e.target.dataset.code;
        copyToClipboard(base64Code);
    } else if (e.target.classList.contains('excel-btn')) {
        const base64Code = e.target.dataset.code;
        applyToExcel(base64Code);
    }
});

// Add helper function to format datetime
function formatDateTime() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

// Initialize Remarkable once at module level
const md = new Remarkable();
// Declare thinkingFlag at global scope
let thinkingFlag = false;

/**
 * chat function to handle user messages, AI streaming, and final rendering.
 * @param {string} role - "User", "AI", or "Error"
 * @param {string} text - The content to display
 * @param {boolean} isFinal - If true, applies Markdown formatting and "Add" buttons
 */
function appendChat(role, text, isFinal = false) {
    const chatHistory = document.getElementById("chat-history");
    let msgDiv = chatHistory.querySelector(".ai.streaming-active");

    // If it's a new message (User, Error, or start of AI response)
    if (!msgDiv || role !== "AI") {
        msgDiv = document.createElement("div");
        msgDiv.className = role.toLowerCase();
        if (role === "AI") msgDiv.classList.add("streaming-active");
        chatHistory.appendChild(msgDiv);
    }

    if (role === "User") {
        // Add datetime timestamp on top of user prompt
        const dateTime = formatDateTime();
        msgDiv.innerHTML = `
            <span class="message-timestamp">${dateTime}</span>
            <strong>You:</strong> ${text}
        `;
    } 
    else if (role === "Error") {
        msgDiv.innerHTML = `<strong>Error:</strong> <span style="color:red;">${text}</span>`;
    } 
    else if (role === "AI") {
        if (!isFinal) {
            // PROGRESSIVE STREAMING: Quick update with simple line breaks
            msgDiv.innerHTML = `<strong>ExcelAI:</strong><br>${text.replace(/\n/g, '<br>')}`;
        } else {
            // FINAL RENDER: Use Remarkable.js for markdown formatting
            msgDiv.classList.remove("streaming-active");
            
            // Custom renderer for code blocks to add buttons
            const originalCodeBlock = md.renderer.rules.code_block;
            const originalFence = md.renderer.rules.fence;

            md.renderer.rules.code_block = function(tokens, idx, options, env, self) {
                const code = tokens[idx].content;
                const lang = tokens[idx].params;
                return createCodeBlockHTML(code, lang);
            };

            md.renderer.rules.fence = function(tokens, idx, options, env, self) {
                const code = tokens[idx].content;
                const lang = tokens[idx].params;
                return createCodeBlockHTML(code, lang);
            };
            
            // Process markdown with Remarkable
            const formattedText = md.render(text);
            msgDiv.innerHTML = `<strong>ExcelAI:</strong><br>${formattedText}`;
        }
    }

    chatHistory.scrollTop = chatHistory.scrollHeight;
    return msgDiv;
}
