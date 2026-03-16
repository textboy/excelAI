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

    const systemInstructions = `
    System Prompt:
    - If response contains any codes or formulas, format them in a markdown code block using three backticks (\`\`\`) instead of single backtick (\`)
    - Priority for solutions: 
    1. Standard Excel UI operations (Menus/Ribbon/Keyboard shortcuts).
    2. Excel Functions/Formulas.
    3. VBA Scripts (only if the above cannot solve it).
    `;
    
    // Combine context and prompt
    const fullPrompt = excelContext 
        ? `Context from Excel sheet:\n${excelContext}\n\nUser Question: ${prompt}\n\n${systemInstructions}`
        : `${prompt}\n\n${systemInstructions}`;

    try {
        // CORRECTED URL
        const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${apiKey}`,
                "Content-Type": "application/json",
                "HTTP-Referer": "http://localhost:3000", // Required by some OpenRouter models
                "X-Title": "Excel AI Bar"
            },
            body: JSON.stringify({
                model: model,
                messages: [{ role: "user", content: fullPrompt }]
            })
        });

        // Check if response is ok
        if (!response.ok) {
            // Try to get detailed error information from the response
            let errorText = `HTTP error! status: ${response.status}`;
            try {
                const errorData = await response.json();
                if (errorData.detail) {
                    errorText += ` - ${errorData.detail}`;
                } else if (errorData.error) {
                    errorText += ` - ${errorData.error}`;
                }
            } catch (e) {
                // If we can't parse JSON, just use the status text
                errorText += ` - ${response.statusText}`;
            }
            throw new Error(errorText);
        }

        const data = await response.json();
        
        if (data.choices && data.choices[0]) {
            appendChat("AI", data.choices[0].message.content);
        } else {
            appendChat("Error", "Unexpected API response format.");
        }
    } catch (error) {
        console.error("Error in sendToAI:", error);
        // Provide more detailed error information
        let errorMessage = `Failed to connect to AI: ${error.message || 'Unknown error'}`;
        
        // If it's a network error, provide additional context
        if (error.name === 'TypeError') {
            errorMessage = "Network error - please check your internet connection and try again.";
        }
        
        appendChat("Error", errorMessage);
    } finally {
        // Re-enable button and restore text
        sendBtn.disabled = false;
        sendBtn.textContent = originalButtonText;
    }
}

function appendChat(role, text) {
    const history = document.getElementById("chat-history");
    const msg = document.createElement("div");
    msg.className = role.toLowerCase();

    // msg.innerHTML = `<strong>${role}:</strong> ${text.replace(/\n/g, '<br>')}`;
    const codeBlockRegex = /```(?:[\w-]+)?\n?([\s\S]*?)```|`([^`\n]+)`/g;
    if (role === "AI") {
        // Regex to find code blocks: ```code```
        let formattedText = text.replace(codeBlockRegex, (match, code) => {
            const cleanCode = code.trim().replace(/"/g, '&quot;').replace(/'/g, '&#39;');
            return `
                <div class="code-block-container">
                    <pre><code>${code.trim()}</code></pre>
                    <button class="apply-btn" onclick="applyToExcel('${cleanCode}')">
                        Add
                    </button>
                </div>`;
        });
        msg.innerHTML = `<strong>AI:</strong><br>${formattedText.replace(/\n/g, '<br>')}`;
    } else {
        msg.innerHTML = `<strong>${role}:</strong> ${text}`;
    }
    
    history.appendChild(msg);
    history.scrollTop = history.scrollHeight;
}

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