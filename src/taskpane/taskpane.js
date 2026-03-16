console.log("taskpane.js loaded");

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

    // --- NEW: FETCH EXCEL CONTEXT ---
    let excelContext = "";
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load("values"); // Load the data from the cells
            await context.sync();
            
            // Convert 2D array of values into a readable string
            const values = range.values;
            excelContext = values.map(row => row.join("\t")).join("\n");
            console.log("--- DEBUG: Excel Context Data ---");
            console.log(excelContext || "[Empty Selection]");
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
    You are an expert Microsoft Excel assistant with deep knowledge of formulas, functions, pivot tables, charts, macros, and data analysis.
    You can provide easy to follow instructions for users.
    You can generate Excel formulas, VBA scripts, and provide instructions for using Excel's features.
    When giving formulas, ensure they are syntactically correct and optimized for performance.
    - If providing an Excel formula, wrap it in \`\`\`excel blocks.
    - If providing VBA code, wrap it in \`\`\`vba blocks.
    - Priority for solutions: 
    1. Standard Excel UI operations (Menus/Ribbon/Keyboard shortcuts).
    2. Excel Functions/Formulas.
    3. VBA Scripts (only if the above cannot solve it).
    `;
    
    // Combine context and prompt
    const fullPrompt = excelContext 
        ? `Context from Excel sheet:\n${excelContext}\n\nUser Question: ${prompt}`
        : `${prompt}`;

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
                    { role: "user", content: fullPrompt }
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

/**
 * Combined chat function to handle user messages, AI streaming, and final rendering.
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
        msgDiv.innerHTML = `<strong>You:</strong> ${text}`;
    } 
    else if (role === "Error") {
        msgDiv.innerHTML = `<strong>Error:</strong> <span style="color:red;">${text}</span>`;
    } 
    else if (role === "AI") {
        if (!isFinal) {
            // PROGRESSIVE STREAMING: Quick update with simple line breaks
            msgDiv.innerHTML = `<strong>AI:</strong><br>${text.replace(/\n/g, '<br>')}`;
        } else {
            // FINAL RENDER: Apply Regex for code blocks and "Add" buttons
            msgDiv.classList.remove("streaming-active");
            
            const codeBlockRegex = /```(?:([\w-]+))?\n?([\s\S]*?)```|`([^`\n]+)`/g;
            
            const formattedText = text.replace(codeBlockRegex, (match, lang, blockCode, inlineCode) => {
                const code = (blockCode || inlineCode).trim();
                const cleanCode = code.replace(/"/g, '&quot;').replace(/'/g, '&#39;');
                const displayLang = lang || (code.startsWith('=') ? 'excel' : 'vba');

                let actionButton = "";
                if (displayLang === "vba") {
                    // VBA Block: Show Copy Button
                    actionButton = `<button class="apply-btn" style="background-color: #666;" onclick="copyToClipboard('${cleanCode}')">Copy VBA</button>`;
                } else {
                    // Excel/Formula Block: Show Add to Excel Button
                    actionButton = `<button class="apply-btn" onclick="applyToExcel('${cleanCode}')">Add to Excel</button>`;
                }

                return `
                    <div class="code-block-container">
                        <div style="font-size:0.7em; color:#aaa; margin-bottom:4px; text-transform:uppercase;">${displayLang}</div>
                        <pre><code>${code}</code></pre>
                        ${actionButton}
                    </div>`;
            });
            
            msgDiv.innerHTML = `<strong>AI:</strong><br>${formattedText.replace(/\n/g, '<br>')}`;
        }
    }

    chatHistory.scrollTop = chatHistory.scrollHeight;
    return msgDiv;
}

/**
 * Copies text content to the system clipboard
 * @param {string} text - The code to copy
 */
async function copyToClipboard(text) {
    try {
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

async function applyToExcel(code) {
    try {
        await Excel.run(async (context) => {
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