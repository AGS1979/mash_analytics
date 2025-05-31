// Declare these globally at the top of your file:
var chatHistory = {};   // Object to store chat history
var currentChat = null; // Current active chat

// Function to initialize a new chat - resets the interface
function resetChatInterface() {
    currentChat = null;
    isFirstQuery = true;
    ticker = null;
    quarter = null;
    year = null;
    askingForKeywords = false;
    document.getElementById("chat-box").innerHTML = '';
    document.getElementById("chat-input").value = '';
    document.getElementById("prev-chat-title").style.display = 'none';
    document.getElementById("new-query-btn").style.display = 'block';
    document.getElementById("send-btn").disabled = false;
    // Clear attached file info if any
    document.getElementById("file-input").value = "";
    document.getElementById("summary-type").style.display = "none";
    document.getElementById("file-info").innerText = "";
}

document.getElementById("summary-type").addEventListener("change", function() {
    const summaryType = this.value;
    const fileInput = document.getElementById("file-input");
    if (summaryType === "esg_compare") {
        // For ESG Comparison, allow multiple HTML files.
        fileInput.accept = "text/html";
        fileInput.multiple = true;
    } else {
        // For other types (including ESG Sustainability Report), accept a single PDF.
        fileInput.accept = "application/pdf";
        fileInput.multiple = false;
    }
});



window.activateChat = function () {
    window.history.pushState({ page: "chat-ui" }, "", "#chat");
    document.getElementById("landing-panel").style.display = "none";
    document.getElementById("chat-ui").style.display = "block";
}

// Handle back button
window.addEventListener("popstate", function (event) {
    document.getElementById("landing-panel").style.display = "block";
    document.getElementById("chat-ui").style.display = "none";
    document.getElementById("custom-agents-ui").style.display = "none"; // ✅ was wrong before

    if (event.state?.page === "chat-ui") {
        document.getElementById("landing-panel").style.display = "none";
        document.getElementById("chat-ui").style.display = "block";
    } else if (event.state?.page === "custom-agents") {
        document.getElementById("landing-panel").style.display = "none";
        document.getElementById("custom-agents-ui").style.display = "block"; // ✅ fixed ID
    }
});

window.useCustomAgents = function () {
    // Hide other sections
    document.getElementById("landing-panel").style.display = "none";
    document.getElementById("chat-ui").style.display = "none";

    // Show the agents panel
    document.getElementById("custom-agents-ui").style.display = "block";

    // Scroll to top of agents section for clean UX
    window.scrollTo({ top: 0, behavior: "smooth" });

    // Push to browser history (so back button works)
    window.history.pushState({ page: "custom-agents" }, "", "#agents");

    // Dynamically load agents if not already loaded
    loadCustomAgents();
}

function showQuantSignalModal() {
    document.getElementById("modal-quant_signal").style.display = "block";
}

function showAgentModal(id) {
    document.getElementById(`modal-${id}`).style.display = "block";
}

function closeModal(id) {
    document.getElementById(`modal-${id}`).style.display = "none";
}

function showPreIPOModal() {
    document.getElementById("modal-Pre-IPO_Investment_Memo").style.display = "block";
}

function showFactorOptModal() {
    document.getElementById("modal-factor_opt").style.display = "block";
}



// ------------------------------------------------------------
    // Definition of loadCustomAgents (this is where two braces were missing)
    // ------------------------------------------------------------
    function loadCustomAgents() {
        const grid = document.getElementById("agent-grid");
        if (!grid) {
            console.error("❌ 'agent-grid' not found!");
            return;
        }

        grid.innerHTML = ''; // Clear old cards

        // Remove old modals
        document.querySelectorAll(".modal").forEach(el => el.remove());

        fetch("/custom-agents-data")
            .then(response => response.json())
            .then(agents => {
                agents.forEach(agent => {
                    // Create the card
                    const card = document.createElement("div");
                    card.className = "agent-card";

                    // 👇 Special handling for Pre-IPO Memo
                    if (agent.id === "Pre-IPO_Investment_Memo") {
                        card.innerHTML = `
                            <h3>${agent.name}</h3>
                            <p><strong>Category:</strong> ${agent.category}</p>
                            <p>${agent.description}</p>
                            <button onclick="showPreIPOModal()">Run Agent</button>
                        `;

                        grid.appendChild(card);  // ✅ ADD THIS

                        const modal = document.createElement("div");
                        modal.id = `modal-${agent.id}`;
                        modal.className = "modal";
                        modal.innerHTML = `
                            <div class="modal-content">
                                <span class="close" onclick="closeModal('${agent.id}')">&times;</span>
                                <h2>${agent.name}</h2>
                                <p><strong>Category:</strong> ${agent.category}</p>
                                <p><strong>Description:</strong> ${agent.description}</p>

                                <form id="preipo-form">
                                    <label>Upload DRHP PDF:</label><br>
                                    <input type="file" name="file" accept=".pdf,.docx" required /><br><br>
                                    <label>Additional Notes / Focus Areas:</label><br>
                                    <textarea name="notes" rows="4" placeholder="e.g. Emphasize risk factors, focus on financials"></textarea><br><br>
                                    <button type="submit">Generate Memo</button>
                                </form>
                                <div id="preipo-result" style="margin-top: 15px;"></div>

                                <hr style="margin: 30px 0;" />
                                <h3>📄 Ask Questions from the Uploaded DRHP</h3>
                                <form id="pdfquery-form">
                                    <label>Ask a Question:</label>
                                    <textarea name="query" rows="3" required placeholder="e.g. What are the key business risks?"></textarea>
                                    <button type="submit">Get Answer</button>
                                </form>
                                <div id="pdfquery-result" style="margin-top: 15px;"></div>
                            </div>
                        `;
                        document.getElementById("custom-agents-ui").appendChild(modal);

                        modal.querySelector("#preipo-form").addEventListener("submit", function (e) {
                            e.preventDefault();
                            const form = e.target;
                            const formData = new FormData(form);

                            const resultDiv = modal.querySelector("#preipo-result");
                            const submitBtn = form.querySelector("button");

                            // Show loading message and disable button
                            resultDiv.innerHTML = `⏳ Generating memo... please wait...`;
                            submitBtn.disabled = true;   

                            fetch("/generate-preipo-memo", {
                                method: "POST",
                                body: formData
                            })
                            .then(r => r.json())
                            .then(data => {
                                const resultDiv = modal.querySelector("#preipo-result");
                                if (data.download_url) {
                                    resultDiv.innerHTML = `
                                        ✅ Memo ready!<br>
                                        <a href="${data.download_url}" target="_blank" style="color:lightblue;font-weight:bold;">
                                            ⬇ Download Memo
                                        </a>`;
                                } else {
                                    resultDiv.innerHTML = `<span style="color:red;">❌ ${data.error}</span>`;
                                }
                            })
                            .catch(err => {
                                console.error(err);
                                modal.querySelector("#preipo-result").innerText = `Error: ${err.message}`;
                            })
                            .finally(() => {
                                submitBtn.disabled = false;
                                resultDiv.scrollIntoView({ behavior: "smooth" });
                            });
                        });

                        modal.querySelector("#pdfquery-form").addEventListener("submit", function (e) {
                            e.preventDefault();
                            const form = e.target;
                            const query = form.query.value.trim();
                            const file = modal.querySelector("input[name='file']").files[0];
                            const resultDiv = modal.querySelector("#pdfquery-result");

                            if (!file || !query) {
                                resultDiv.innerHTML = "Please upload a DRHP and enter a query.";
                                return;
                            }

                            const formData = new FormData();
                            formData.append("file", file);
                            formData.append("query", query);

                            resultDiv.innerHTML = "⏳ Querying document...";
                            fetch("/query-pdf", {
                                method: "POST",
                                body: formData
                            })
                            .then(r => r.json())
                            .then(data => {
                                if (data.answer) {
                                    resultDiv.innerHTML = `
                                        <div style="margin-top:10px;" class="pdf-answer-box">
                                            <div>${data.answer}</div>
                                            <div style="margin-top: 8px;">
                                                <small><em>Cited pages:</em> ${data.pages.join(', ')}</small>
                                            </div>
                                        </div>
                                    `;
                                } else {
                                    resultDiv.innerHTML = `<span style="color:red;">❌ ${data.error || 'No response'}</span>`;
                                }
                            })
                            .catch(err => {
                                console.error(err);
                                resultDiv.innerText = `Error: ${err.message}`;
                            });
                        });

                    } else if (agent.id === "quant_signal") {
  // 1) Create the card (same as before)
  card.innerHTML = `
    <h3>${agent.name}</h3>
    <p><strong>Category:</strong> ${agent.category}</p>
    <p>${agent.description}</p>
    <button onclick="showQuantSignalModal()">Run Agent</button>
  `;
  grid.appendChild(card);

  // 2) Create the modal container
  const modal = document.createElement("div");
  modal.id = `modal-${agent.id}`;
  modal.className = "modal";
  modal.innerHTML = `
    <div class="modal-content">
      <span class="close" onclick="closeModal('${agent.id}')">&times;</span>
      <h2>${agent.name}</h2>
      <p><strong>Category:</strong> ${agent.category}</p>
      <p><strong>Description:</strong> ${agent.description}</p>

      <!-- FORM STARTS HERE -->
      <form id="quant-signal-form" enctype="multipart/form-data">
        <label>Upload Deal PDF:</label><br>
        <input type="file" name="file" accept=".pdf" required /><br><br>

        <label>What do you want to extract (e.g., valuation, SWOT, red flags)?</label><br>
        <textarea name="query" rows="3" required placeholder="e.g. valuation, SWOT analysis, red flags"></textarea><br><br>

        <!-- NEW: Deep‐Dive Master Toggle -->
        <label>
          <input type="checkbox" name="run_deep_dives" value="yes" checked />
          Run Detailed “Deep‐Dive” Analysis?
        </label><br><br>

        <!-- NEW: Individual Deep‐Dive Sections -->
        <div style="margin-left: 16px; margin-bottom: 12px;">
          <label><input type="checkbox" name="sections" value="Market Analysis" checked /> Market Analysis</label><br>
          <label><input type="checkbox" name="sections" value="Financial Performance" checked /> Financial Performance</label><br>
          <label><input type="checkbox" name="sections" value="Operational Risks" checked /> Operational Risks</label><br>
          <label><input type="checkbox" name="sections" value="Exit Strategy" checked /> Exit Strategy</label><br>
        </div>

        <button type="submit">Run Analysis</button>
      </form>
      <!-- FORM ENDS HERE -->

      <hr style="margin:24px 0;" />

      <div id="quant-signal-result" style="max-height: 60vh; overflow-y: auto;"></div>
    </div>
  `;
  document.getElementById("custom-agents-ui").appendChild(modal);

  // 3) Handle form submission
  modal.querySelector("#quant-signal-form").addEventListener("submit", function (e) {
    e.preventDefault();
    const form = e.target;
    const fileInput = form.querySelector("input[name='file']");
    const queryText = form.querySelector("textarea[name='query']").value.trim();
    const resultDiv = document.getElementById("quant-signal-result");
    const submitBtn = form.querySelector("button");

    if (!fileInput.files[0] || !queryText) {
      resultDiv.innerHTML = "<p style='color: red;'>Please upload a PDF and enter a query.</p>";
      return;
    }

    // Build FormData payload
    const formData = new FormData();
    formData.append("file", fileInput.files[0]);
    formData.append("query", queryText);

    // Whether to run deep dives at all?
    const runDeepDives = form.querySelector("input[name='run_deep_dives']").checked;
    formData.append("run_deep_dives", runDeepDives ? "true" : "false");

    // Which sections to deep dive?
    // We allow multiple checkboxes with name="sections"
    const selectedSections = Array.from(
      form.querySelectorAll("input[name='sections']:checked")
    ).map((el) => el.value);
    // Convert the array to JSON‐string so Flask can parse
    formData.append("sections", JSON.stringify(selectedSections));

    // Show “loading” message
    resultDiv.innerHTML = "<p>⏳ Running analysis, please wait...</p>";
    submitBtn.disabled = true;

    // 4) POST to /analyze_deal (instead of /process_report)
    fetch("/analyze_deal", {
      method: "POST",
      body: formData
    })
      .then(async (resp) => {
        if (!resp.ok) {
          const err = await resp.json();
          throw new Error(err.error || `HTTP ${resp.status}`);
        }
        return resp.json();
      })
      .then((data) => {
        // 5) Render the returned JSON: data.aggregate_analysis + deep dives
        let html = "";

        // (A) Render the aggregated summary JSON as formatted HTML
        if (data.aggregate_analysis) {
          html += "<h3>📦 Aggregate Analysis</h3>";
          // If the backend returned a JSON object for “aggregate_analysis”:
          try {
            const agg = data.aggregate_analysis;
            // If it’s literally an object with keys, stringify nicely:
            if (typeof agg === "object") {
              html += "<pre style='background:#222; padding:12px; border-radius:6px;'>";
              html += JSON.stringify(agg, null, 2);
              html += "</pre>";
            } else {
              // If it’s just raw text fallback
              html += `<div>${agg}</div>`;
            }
          } catch (err) {
            html += `<div>Unable to parse aggregate_analysis.</div>`;
          }
        }

        // (B) If deep dive sections exist, render each block
        if (runDeepDives && Array.isArray(selectedSections)) {
          selectedSections.forEach((sectionName) => {
            const key = `deep_dive_${sectionName.replace(/ /g, "_")}`;
            if (data[key]) {
              html += `<hr style="margin:16px 0;" />`;
              html += `<h3>🔍 Deep Dive: ${sectionName}</h3>`;
              // The backend returned Markdown (or raw HTML). We’ll assume it’s text—just wrap in a <pre>.
              html += `<pre style="background:#222; padding:12px; border-radius:6px;">`;
              html += data[key];
              html += "</pre>";
            }
          });
        }

        resultDiv.innerHTML = html || "<p>No analysis returned.</p>";
      })
      .catch((err) => {
        console.error(err);
        resultDiv.innerHTML = `<p style="color:red;">❌ Error: ${err.message}</p>`;
      })
      .finally(() => {
        submitBtn.disabled = false;
        resultDiv.scrollIntoView({ behavior: "smooth" });
      });
  });
}
 else if (agent.id === "factor_opt") {
                        card.innerHTML = `
                            <h3>${agent.name}</h3>
                            <p><strong>Category:</strong> ${agent.category}</p>
                            <p>${agent.description}</p>
                            <button onclick="showFactorOptModal()">Run Agent</button>
                        `;
                        grid.appendChild(card);  // ✅ ADD THIS

                        const modal = document.createElement("div");
                        modal.id = `modal-${agent.id}`;
                        modal.className = "modal";
                        modal.innerHTML = `
                            <div class="modal-content">
                                <span class="close" onclick="closeModal('${agent.id}')">&times;</span>
                                <h2>${agent.name}</h2>
                                <p><strong>Category:</strong> ${agent.category}</p>
                                <p><strong>Description:</strong> ${agent.description}</p>
                                <form id="factoropt-form">
                                    <label>Upload Portfolio CSV:</label><br>
                                    <input type="file" name="file" accept=".csv" required /><br><br>
                                    <label>Target Factor Exposures (JSON):</label><br>
                                    <input type="text" name="target_exposures" placeholder='{"MKT":0.1,"SMB":0.05,"HML":0.02, "RMW": 0.01, "CMA": 0.00}' required><br><br>
                                    <label>Turnover Limit (Optional):</label><br>
                                    <input type="text" name="turnover_limit" placeholder="0.1"><br><br>
                                    <button type="submit">Optimize Portfolio</button>
                                </form>
                                <div id="factoropt-result" style="margin-top: 15px;"></div>
                            </div>
                        `;
                        document.getElementById("custom-agents-ui").appendChild(modal);

                        modal.querySelector("form").addEventListener("submit", function (e) {
                            e.preventDefault();
                            const form = e.target;
                            const formData = new FormData(form);
                            const resultDiv = modal.querySelector("#factoropt-result");
                            const submitBtn = form.querySelector("button");

                            resultDiv.innerHTML = `⏳ Running optimizer...`;
                            submitBtn.disabled = true;

                            fetch("/optimize-factor-portfolio", {
                                method: "POST",
                                body: formData
                            })
                            .then(r => r.json())
                            .then(data => {
                                if (data.status === "optimal") {
                                    const weights = data.optimized_weights.map(w => `<li>${w.ticker}: ${w.weight}</li>`).join('');
                                    const exposures = Object.entries(data.achieved_exposures).map(
                                        ([k, v]) => `${k}: ${v}`
                                    ).join("<br>");
                                    resultDiv.innerHTML = `
                                        ✅ Optimization Complete<br>
                                        <strong>New Weights:</strong><br><ul>${weights}</ul>
                                        <strong>Achieved Exposures:</strong><br>${exposures}
                                    `;
                                } else {
                                    resultDiv.innerHTML = `<span style="color:red;">❌ ${data.message || 'Optimization failed.'}</span>`;
                                }
                            })
                            .catch(err => {
                                console.error(err);
                                resultDiv.innerText = `Error: ${err.message}`;
                            })
                            .finally(() => {
                                submitBtn.disabled = false;
                                resultDiv.scrollIntoView({ behavior: "smooth" });
                            });
                        });
                    } else {
                        // Default cards for other agents
                        card.innerHTML = `
                            <h3>${agent.name}</h3>
                            <p><strong>Category:</strong> ${agent.category}</p>
                            <p>${agent.description}</p>
                            <button onclick="showAgentModal('${agent.id}')">Learn More</button>
                        `;
                        grid.appendChild(card);  // ✅ ADD THIS
                        const modal = document.createElement("div");
                        modal.id = `modal-${agent.id}`;
                        modal.className = "modal";
                        modal.innerHTML = `
                            <div class="modal-content">
                                <span class="close" onclick="closeModal('${agent.id}')">&times;</span>
                                <h2>${agent.name}</h2>
                                <p><strong>Category:</strong> ${agent.category}</p>
                                <p><strong>How it Works:</strong> ${agent.description}</p>
                                <p><strong>Sample Output:</strong></p>
                                <pre>${agent.output || "No sample output available."}</pre>
                                <a href="/static/samples/${agent.id}_output.csv" download class="download-link">⬇ Download Sample Output</a>
                            </div>
                        `;
                        document.getElementById("custom-agents-ui").appendChild(modal);
                    }
                }); // ← Closes agents.forEach(agent => { … })

            } // <-- added: closes the `.then(agents => { … })` callback
        );   // completes `fetch(...).then(...).then(...);`
    }      // <-- added: closes function loadCustomAgents()

    // ------------------------------------------------------------
    // (End of loadCustomAgents)
    // ------------------------------------------------------------

document.addEventListener("DOMContentLoaded", function () {
    
    // Immediately set the current chat to the logged-in user if available.
    if (typeof currentUser !== "undefined" && currentUser) {
        currentChat = currentUser;  // Use the username as the key for this chat session
        loadUserChat(currentUser);
    }

    const hash = window.location.hash;

    // Default state: show landing panel only
    document.getElementById("landing-panel").style.display = "block";
    document.getElementById("chat-ui").style.display = "none";
    document.getElementById("custom-agents-ui").style.display = "none";

    if (hash === "#chat") {
        activateChat();
    } else if (hash === "#agents") {
        useCustomAgents();
    }




    // After loadCustomAgents, the code continues as before...
    const name = sessionStorage.getItem("userFirstName");
    const company = sessionStorage.getItem("userCompany");

    if (name && company) {
        const nameEl = document.getElementById("user-name-span");
        const companyEl = document.getElementById("user-company-span");

        if (nameEl && companyEl) {
            nameEl.textContent = name;
            companyEl.textContent = company;
        }
    }

    let isFirstQuery = true; // Flag to track if it's the first query in a new chat
    let ticker = null;
    let quarter = null;
    let year = null;
    let askingForKeywords = false; // Flag to track if we're asking for keywords

    // Function to create a new chat entry in history
    function createNewChat() {
        const chatCount = Object.keys(chatHistory).length + 1;
        currentChat = 'Chat' + chatCount;
        chatHistory[currentChat] = [];
        updateChatHistoryList();
    }

    // -------------------------
    // New Helper Functions for 10K Queries
    // -------------------------

    // Detect if the query is about 10K filings
    function is10KQuery(query) {
        return query.toLowerCase().includes("10k");
    }

    // Extract parameters from queries like:
    // "Download 10K for Apple for 2024" or "generate 10Ks for Apple for 2021,2022"
    function extract10KParameters(query) {
        // Using a regex for robustness (case-insensitive)
        let regex = /10k\s+for\s+(.+?)\s+for\s+([\d,]+)/i;
        let match = query.match(regex);
        if (match) {
            return { company_name: match[1].trim(), years: match[2].trim() };
        }
        return null;
    }

    // Detect if the user is asking for a management-evasiveness analysis
    function isEvasivenessQuery(q) {
        q = q.toLowerCase();
        return (
            /evasive/.test(q) ||
            (/\b(evasiveness|analysis)\b/.test(q) && /\b(call|quarter)\b/.test(q))
        );
    }

    // Parse out "company", "quarter" and "year" from queries like
    //   "Evasiveness analysis for Boeing Q3 2024"
    function extractEvasivenessParameters(query) {
        const q = query.toLowerCase();
        
        let z;

        // 0) "How evasive was Tesla in Q1 2025"
        let re0 = /(?:how\s+evasive\s+(?:was|is)\s+)(.+?)\s+in(?:\s+\w+){0,3}?\s+q([1-4])\s+(\d{4})/i;
        if (z = query.match(re0)) {
            return {
                company: z[1].trim(),
                quarter: parseInt(z[2],10),
                year:    parseInt(z[3],10)
            };
        }

        // 1) old style: "... for X Q2 2024"
        let re1 = /\bfor\s+(.+?)\s+q([1-4])\s+(\d{4})/i;

        // 2) "X's fourth quarter of 2024"
        let re2 = /(?:for|on)?\s*([\w &]+?)'s?\s*(?:\b(?:1st|2nd|3rd|4th|first|second|third|fourth)\b)\s*quarter\s*(?:of)?\s*(\d{4})/i;

        // 3) "quarter 4 2024 for X"
        let re3 = /quarter\s*([1-4])\s*(?:of)?\s*(\d{4}).*?\bfor\s+(.+?)$/i;

        let m;
        if (m = query.match(re1)) {
            return { company: m[1].trim(), quarter: +m[2], year: +m[3] };
        }
        if (m = query.match(re2)) {
            // re2 captures [company, year], but we need to infer quarter from the ordinal
            let ord = query.match(/\b(1st|2nd|3rd|4th|first|second|third|fourth)\b/i)[1].toLowerCase();
            let qmap = { first:1, "1st":1, second:2, "2nd":2, third:3, "3rd":3, fourth:4, "4th":4 };
            return { company: m[1].trim(), quarter: qmap[ord], year: +m[2] };
        }
        if (m = query.match(re3)) {
            return { company: m[3].trim(), quarter: +m[1], year: +m[2] };
        }
        return null;
    }

    // Handle the API call for 10K queries by fetching the binary response and triggering a download.
    function handle10KApiCall(endpoint, body, chatBox) {
        fetch(endpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(body)
        })
        .then(function(response) {
            if (!response.ok) {
                throw new Error("Server responded with status " + response.status);
            }
            return response.blob();
        })
        .then(function(blob) {
            // Create a temporary download link for the file (PDF or ZIP)
            var url = window.URL.createObjectURL(blob);
            var a = document.createElement("a");
            a.href = url;
            a.download = "10K_filings_download"; // Default filename
            document.body.appendChild(a);
            a.click();
            a.remove();
            addChatMessage("bot", "Your 10K filings are ready for download.");
        })
        .catch(function(error) {
            console.error("Error in 10K API call:", error);
            addChatMessage("bot", "An error occurred: " + error.message);
        });
    }

    function loadUserChat(username) {
        fetch(`/load_chat?username=${encodeURIComponent(username)}`)
            .then(response => response.json())
            .then(data => {
                if (data.chat && data.chat.length > 0) {
                    chatHistory[username] = data.chat;
                    // Display the chat history in the chat box
                    document.getElementById("chat-box").innerHTML = "";
                    data.chat.forEach(entry => {
                        document.getElementById("chat-box").innerHTML += `<div class="${entry.type}-message">${entry.message}</div>`;
                    });
                }
            })
            .catch(err => console.error("Error loading chat:", err));
    }

    // Update the chat history list in the left column
    function updateChatHistoryList() {
        const chatHistoryList = document.getElementById("chat-history-list");
        chatHistoryList.innerHTML = '';
        const previewLimit = 80;
        for (let chat in chatHistory) {
            let chatItem = document.createElement("li");
            let tempDiv = document.createElement("div");
            tempDiv.innerHTML = chatHistory[chat].length > 0 ? chatHistory[chat][0].message : "Empty Chat";
            const fullText = tempDiv.innerText || tempDiv.textContent;
            const truncatedText = fullText.length > previewLimit ? fullText.substring(0, previewLimit) + "..." : fullText;
            chatItem.innerText = `${chat}: ${truncatedText}`;
            chatItem.onclick = function() {
                viewChatHistory(chat);
            };
            chatHistoryList.appendChild(chatItem);
        }
    }

    // Function to view selected chat history
    function viewChatHistory(chatName) {
        currentChat = chatName;
        const chatBox = document.getElementById("chat-box");
        chatBox.innerHTML = '';
        isFirstQuery = false;
        chatHistory[chatName].forEach(entry => {
            if (entry.type === 'bot') {
                chatBox.innerHTML += entry.message;
            } else {
                chatBox.innerHTML += `<div class='user-message'>${entry.message}</div>`;
            }
        });
    }

    // Standard API call handler for non-file endpoints
    function handleApiCall(endpoint, body, chatBox) {
        console.log("Initial payload sent to API:", body);
        console.log("Sending request to:", endpoint);
        console.log("Request body:", JSON.stringify(body));

        fetch(endpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(body),
        })
        .then((response) => {
            if (!response.ok) {
                throw new Error(`Server responded with status ${response.status}`);
            }
            return response.json();
        })
        .then((data) => {
            console.log("📢 API Response:", data);

            let botMessageHtml = "";
            if (data.download_url) {
                botMessageHtml = `
                    <div class='bot-message'>
                        ✅ Report is ready!<br>
                        <a href="${data.download_url}"
                            target="_blank"
                            style="color:blue;font-weight:bold;">
                            ⬇ Download Report
                        </a>
                    </div>
                `;
            }
            // Handle Stock Report Response
            else if (data.file_url) {
                botMessageHtml = `<div class='bot-message'>
                    ✅ Stock report generated successfully! <br>
                    <a href="${data.file_url}" target="_blank" style="color: blue; font-weight: bold;">
                        ⬇ Download Report
                    </a>
                </div>`;
            }
            // Handle News Response
            else if (data.news && Array.isArray(data.news)) {
                if (data.news.length > 0) {
                    botMessageHtml += `<div class='bot-message'><strong>Latest News Updates:</strong></div>`;
                    data.news.forEach(article => {
                        botMessageHtml += `<div class='bot-message'>
                            <strong>${article.title}</strong><br>
                            ${article.summary}<br>
                            <a href="${article.url}" target="_blank">Read More</a>
                        </div>`;
                    });
                } else {
                    botMessageHtml += `<div class='bot-message'>No news articles found.</div>`;
                }
            }
            // Handle Document Processing Response
            else if (data.document_url) {
                let message = data.message ? data.message : "Document processed successfully!";
                botMessageHtml += `<div class='bot-message'>
                    <strong>${message}</strong><br>
                    <a href="${data.document_url}" target="_blank">Download Generated Document</a>
                </div>`;
            }
            // Handle numerical sentence extraction (figures from earnings call)
            else if (data.company && data.quarter && data.year && data.sentences) {
                console.log("📊 API Response:", data);

                let botMessageHtml = `<div class='bot-message'><strong>📊 Extracted Figures from Earnings Call:</strong><br>`;
                data.sentences.slice(0, 15).forEach(sentence => {
                    botMessageHtml += `${sentence}<br>`;
                });
                botMessageHtml += `...and ${data.sentences.length - 15} more sentences. `;
                botMessageHtml += `<br><a href="/download/${data.company}_${data.year}_Q${data.quarter}_figures.xlsx" target="_blank" style="color:blue;font-weight:bold;">⬇ Download Full Report</a></div>`;

                const wrapper = document.createElement("div");
                wrapper.className = "bot-message";
                wrapper.innerHTML = botMessageHtml;
                chatBox.appendChild(wrapper);
                chatHistory[currentChat].push({ type: 'bot', message: botMessageHtml });
                chatBox.scrollTop = chatBox.scrollHeight;
                return;
            }

            // Handle Earnings Call Summary
            else if (data.summary) {
                botMessageHtml += `<div class='bot-message'><strong>Earnings Call Summary:</strong><br>${data.summary}</div>`;
            }
            else if (data.red_flags && Array.isArray(data.red_flags)) {
                let botMessageHtml = `<div class='bot-message'><strong>🚩 Red Flag Statements Detected (${data.count}):</strong><br>`;
                data.red_flags.forEach((entry, idx) => {
                    botMessageHtml += `<span><strong>[${entry.score}]</strong> ${entry.sentence}</span><br>`;
                });
                botMessageHtml += `</div>`;
                const wrapper = document.createElement("div");
                wrapper.className = "bot-message";
                wrapper.innerHTML = botMessageHtml;
                chatBox.appendChild(wrapper);
                chatHistory[currentChat].push({ type: 'bot', message: botMessageHtml });
                chatBox.scrollTop = chatBox.scrollHeight;
                return;
            }

            // Handle General Message Responses
            else if (data.response) {
                botMessageHtml = `<div class='bot-message'>${data.response}</div>`;
            }
            // Handle Guidance Change (very likely just "message" key)
            else if (endpoint === '/analyze-guidance-change' && data.message) {
                botMessageHtml = `
                    <div class='bot-message'>
                        <strong>🔎 Guidance Update Result:</strong><br>
                        ${data.message}
                    </div>
                `;
            }
            // General fallback for message responses
            else if (data.message) {
                botMessageHtml = `<div class='bot-message'>${data.message}</div>`;
            }
            // Handle Unexpected Response Formats
            else {
                console.error("Unexpected response format:", data);
                botMessageHtml = `<div class='bot-message'>⚠ Unexpected response format.</div>`;
            }

            const wrapper = document.createElement("div");
            wrapper.className = "bot-message";
            wrapper.innerHTML = botMessageHtml;
            chatBox.appendChild(wrapper);

            chatHistory[currentChat].push({ type: 'bot', message: botMessageHtml });
            chatBox.scrollTop = chatBox.scrollHeight;
        })
        .catch((error) => {
            console.error("❌ Error during API call:", error);
            const errorMsg = `<div class='bot-message'>An error occurred: ${error.message}</div>`;
            chatBox.innerHTML += errorMsg;
            chatHistory[currentChat].push({ type: 'bot', message: errorMsg });
            chatBox.scrollTop = chatBox.scrollHeight;
        });
    }

    // Helper function to detect if the user query is asking for a PDF summary.
    function isPdfSummaryQuery(query) {
        const phrases = ["pdf summary", "upload pdf", "document summary"];
        return phrases.some(phrase => query.toLowerCase().includes(phrase));
    }

    // Helper function to detect if the user query is for an earnings call summary.
    function isEarningsCallSummaryQuery(query) {
        const q = query.toLowerCase();
        return (
            q.includes("summary of earnings call") ||
            q.includes("summarize earnings call") ||
            q.includes("can you summarize earnings call") ||
            (q.includes("get me a summary") && q.includes("earnings") && q.includes("call")) ||
            q.match(/q[1-4]\s+\d{4}.*earnings call summary/) ||
            q.match(/earnings call.*(overview|recap|summary)/) ||
            q.match(/summarize\s+.+?'s\s+q[1-4]\s+\d{4}\s+earnings call/i) ||  // ⬅ handles "Caterpillar's Q1 2025 earnings call"
            q.match(/can you summarize\s+.+?'s\s+q[1-4]\s+\d{4}\s+earnings call/i)
        );
    }

    // Attach button event to trigger file input click
    document.getElementById("attach-btn").addEventListener("click", function() {
        document.getElementById("file-input").click();
    });

    // When a file is selected, show file info and summary type dropdown
    document.getElementById("file-input").addEventListener("change", function() {
        const fileInput = document.getElementById("file-input");
        const fileInfo = document.getElementById("file-info");

        if (fileInput.files.length > 0) {
            let fileNames = Array.from(fileInput.files).map(file => file.name).join(", ");
            fileInfo.innerText = `Attached: ${fileNames}`;
            document.getElementById("summary-type").style.display = "inline-block";
        } else {
            fileInfo.innerText = "";
            document.getElementById("summary-type").style.display = "none";
        }
    });

    // Add a helper function to append a chat message
    function addChatMessage(type, message) {
        const chatBox = document.getElementById("chat-box");
        const msgHtml = `<div class='${type === "bot" ? "bot-message" : "user-message"}'>${message}</div>`;
        chatBox.innerHTML += msgHtml;
        chatHistory[currentChat] = chatHistory[currentChat] || [];
        chatHistory[currentChat].push({ type: type, message: message });
        chatBox.scrollTop = chatBox.scrollHeight;
        fetch('/save_chat', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ username: currentChat, chat: chatHistory[currentChat] })
        })
        .then(response => response.json())
        .then(data => console.log("Chat saved", data))
        .catch(err => console.error("Error saving chat:", err));
    }

    // Updated send button event handler for file-based queries
    document.getElementById("send-btn").addEventListener("click", function () {
        const chatBox   = document.getElementById("chat-box");
        const fileInput = document.getElementById("file-input");

        // 1) File-upload branch
        if (fileInput.files.length > 0) {
            if (!chatHistory[currentChat]) createNewChat();
            const summaryType = document.getElementById("summary-type").value;
            const queryText   = document.getElementById("chat-input").value.trim();

            if (!queryText) {
                addChatMessage("bot", "Please enter a query along with your file upload.");
                return;
            }
            addChatMessage("user", queryText);

            const formData = new FormData();
            formData.append('query', queryText);
            let endpoint;

            if (summaryType === 'short') {
                formData.append('file', fileInput.files[0]);
                endpoint = '/document-short-summary';
            } else if (summaryType === 'long') {
                formData.append('file', fileInput.files[0]);
                endpoint = '/document-long-summary';
            } else if (summaryType === 'esg') {
                formData.append('file', fileInput.files[0]);
                endpoint = '/process_report';
            } else if (summaryType === 'esg_compare') {
                for (let i = 0; i < fileInput.files.length; i++) {
                    formData.append('files', fileInput.files[i]);
                }
                endpoint = '/compare_insights';
            } else if (summaryType === 'earnings-call') {
                formData.append('file', fileInput.files[0]);
                endpoint = '/generate-earnings-call-summary';
            } else if (summaryType === 'redflag') {
                formData.append('file', fileInput.files[0]);
                endpoint = '/analyze-redflags';
            }

            fetch(endpoint, { method: 'POST', body: formData })
            .then(r => r.json())
            .then(data => {
                if (data.error) {
                    addChatMessage("bot", `Error: ${data.error}`);
                    return;
                }

                // ✅ Handle inline red flag display for file-based input
                if (data.red_flags && Array.isArray(data.red_flags)) {
                    let html = `<div class='bot-message'><strong>🚩 Red Flag Statements Detected (${data.count}):</strong><br>`;
                    data.red_flags.forEach(entry => {
                        html += `<span><strong>[${entry.score}]</strong> ${entry.sentence}</span><br>`;
                    });
                    html += `</div>`;
                    addChatMessage("bot", html);
                    return;
                }

                // fallback in case only download URL was provided
                if (data.download_url) {
                    let botHtml = `<div class='bot-message'>
                        ✅ Report Ready!<br>
                        <a href="${data.download_url}" target="_blank" style="color:blue;font-weight:bold;">
                            ⬇ Download Report
                        </a></div>`;
                    addChatMessage("bot", botHtml);
                    return;
                }

                // generic message fallback
                addChatMessage("bot", `<div class='bot-message'>✅ Report Ready!</div>`);
            })

            .catch(err => {
                console.error(err);
                addChatMessage("bot", `An error occurred: ${err.message}`);
            })
            .finally(() => {
                fileInput.value = "";
                document.getElementById("summary-type").style.display = "none";
                document.getElementById("file-info").innerText = "";
                document.getElementById("chat-input").value = "";
            });

            return; // stop here if file upload was handled
        }

        // 2) Text-only branch
        const query = document.getElementById("chat-input").value.trim();
        if (!query) return;
        if (!chatHistory[currentChat]) createNewChat();
        document.getElementById("chat-input").value = '';
        addChatMessage("user", query);
        updateChatHistoryList();

        // 2a) Management Evasiveness first
        if (isEvasivenessQuery(query)) {
            const params = extractEvasivenessParameters(query);
            // if we could parse out explicit fields, send them…
            if (params) {
                handleApiCall('/analyze-evasiveness', params, chatBox);
            }
            // otherwise, send the raw text as `{ query: … }`
            else {
                handleApiCall('/analyze-evasiveness', { query }, chatBox);
            }
            return;
        }

        // ✅ Add here
        if (isRedFlagQuery(query)) {
            handleApiCall('/analyze-redflags', { text: query }, chatBox);
            return;
        }

        // 2b) 10-K Filings
        if (is10KQuery(query)) {
            const params = extract10KParameters(query);
            if (!params) {
                addChatMessage("bot",
                    "Use: 'generate 10Ks for <Company> for <Years>'."
                );
                return;
            }
            handle10KApiCall('/get10k', params, chatBox);
            return;
        }

        // 2c) Other text queries
        else if (isPdfSummaryQuery(query)) {
            addChatMessage("bot", "Please attach a PDF using the '+' button.");
        } else if (isEarningsCallSummaryQuery(query)) {
            handleApiCall('/generate-earnings-call-summary', { message: query }, chatBox);
        } else if (isEarningsReportQuery(query)) {
            handleApiCall('/generate-earnings-report', { query }, chatBox);
        } else if (isNewsQuery(query)) {
            handleApiCall('/fetch-news-updates', { query }, chatBox);
        } else if (isMarketPerformanceQuery(query)) {
            handleApiCall('/generate-market-performance', { query }, chatBox);
        } else if (isDebtCovenantQuery(query)) {
            handleApiCall('/generate-debt-covenants', { query }, chatBox);
        } else if (isGuidanceChangeQuery(query)) {
            handleApiCall('/analyze-guidance-change', { query }, chatBox);
        } else if (isStockReportQuery(query)) {
            handleApiCall('/generate-report', { query }, chatBox);
        } else if (isEarningsCallExtractionQuery(query)) {
            handleApiCall('/extract-earnings-call-data', { query }, chatBox);
        } else {
            addChatMessage("bot", "Your query was received. (Other queries are handled elsewhere.)");
        }
    }); // ← end of send-btn listener


    // New Query Button to Reset
    document.getElementById("new-query-btn").addEventListener("click", function() {
        resetChatInterface();
    });



    function isRedFlagQuery(query) {
        return query.toLowerCase().includes("red flag") || query.toLowerCase().includes("flagged statement");
    }

    function isEarningsCallExtractionQuery(query) {
        const patterns = [
            /\bextract (the )?figures\b/i,
            /\bget (the )?figures\b/i,
            /\bnumerical figures\b/i,
            /\bextract numbers from\b/i,
            /\bnumbers from\b/i,
            /\bmetrics from\b/i,
            /\bhighlight figures\b/i,
            /\bextract data\b/i,
            /\bextract the data from\b/i,
            /\bdata from earnings call\b/i,
            /\bfigures from earnings call\b/i
        ];

        return patterns.some(pattern => pattern.test(query)) && query.toLowerCase().includes("earnings call");
    }

    function isEarningsReportQuery(query) {
        const q = query.toLowerCase();
        return (
            q.includes("extract keywords from") ||
            q.includes("earnings report with") ||
            q.includes("analyze earnings call for") ||
            (q.includes("keywords") && q.includes("earnings call")) ||
            (q.includes("growth") || q.includes("outlook") || q.includes("margin"))
        );
    }

    function isDebtCovenantQuery(query) {
        const phrases = [
            "debt covenant",
            "extract debt terms",
            "loan covenant",
            "financial covenants",
            "bond covenant",
            "credit agreement terms"
        ];
        return phrases.some(phrase => query.toLowerCase().includes(phrase));
    }


    function isGuidanceChangeQuery(query) {
        const phrases = [
            "guidance", "upgrade guidance", "downgrade guidance",
            "raise forecast", "lower forecast", "guidance change",
            "was guidance upgraded", "was guidance downgraded",
            "full year guidance", "revised guidance"
        ];
        return phrases.some(phrase => query.toLowerCase().includes(phrase));
    }

    function isMarketPerformanceQuery(query) {
        const phrases = [
            "global market performance",
            "emerging market performance",
            "developed market performance",
            "market performance dashboard",
            "global stock markets",
            "emerging equity markets",
            "developed equity markets",
            "global performance report",
            "emerging performance report",
            "developed performance report",
            "all market performance",
            "advanced markets",
            "developing markets"
        ];
        return phrases.some(phrase => query.toLowerCase().includes(phrase));
    }

    function isNewsQuery(query) {
        const phrases = [
            "latest news",
            "recent updates",
            "what's happening",
            "tell me",
            "latest update",
            "news",
            "sector update",
            "industry trends",
            "market news",
            "latest in",
            "what is the latest news",
            "what are the recent updates",
            "give me a sector update",
            "what are the industry trends"
        ];
        return phrases.some(phrase => query.toLowerCase().includes(phrase));
    }

    function isStockReportQuery(query) {
        const phrases = [
            "stock report",
            "stock note",
            "company report",
            "company note",
            "short report",
            "short note",
            "equity report",
            "equity note"
        ];
        return phrases.some(phrase => query.toLowerCase().includes(phrase));
    }

    function cleanAndValidateKeywords(input) {
        return input.split(',').map(keyword => keyword.trim()).filter(keyword => keyword.length > 0);
    }

}); // ✅ Added closing bracket for DOMContentLoaded event listener
