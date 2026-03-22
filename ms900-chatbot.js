/* ─── MS-900 AI Study Assistant ─────────────────────────────────────────────
   Floating chatbot widget powered by WebLLM (runs offline in your browser).
   - No API key required — the model runs locally using WebGPU
   - Model is cached after first download (no re-download on page reload)
   - Extracts current page content for context-aware answers
   - Includes condensed MS-900 knowledge base for cross-page questions
   ──────────────────────────────────────────────────────────────────────────── */

(function() {
  'use strict';

  const HISTORY_KEY = 'ms900_chatbot_history';
  const MAX_HISTORY = 10; // max message pairs to keep in conversation
  const MODEL_ID_F16 = 'Llama-3.2-1B-Instruct-q4f16_1-MLC'; // ~800MB, requires GPU f16 support
  const MODEL_ID_F32 = 'Llama-3.2-1B-Instruct-q4f32_1-MLC'; // ~800MB, wider GPU compatibility (mobile)

  // ─── KNOWLEDGE BASE (condensed from all training pages) ──────────────────
  const KNOWLEDGE_BASE = `
You are an MS-900 (Microsoft 365 Fundamentals) study assistant. Answer questions ONLY about MS-900 exam topics using the knowledge below. Be concise, accurate, and exam-focused. When relevant, mention exam tips. If a question is outside MS-900 scope, say so politely.

KEY FACTS BY DOMAIN:

DOMAIN 1 — CLOUD CONCEPTS (5-10% of exam):
- Cloud computing = renting compute/storage/software over the internet instead of owning hardware. OpEx (pay-as-you-go) vs CapEx (upfront purchase).
- Shared Responsibility Model: Physical datacenter=always Microsoft (IaaS/PaaS/SaaS). OS/VMs=Customer in IaaS, Microsoft in PaaS/SaaS. Applications=Customer in IaaS/PaaS, Microsoft in SaaS. Data & Identities=ALWAYS Customer in all models.
- Deployment Models: Public (shared, scalable, Azure), Private (single org, full control), Hybrid (both connected).
- Service Models: IaaS (you manage OS+apps, e.g. Azure VMs), PaaS (you manage code+data, e.g. Azure App Service), SaaS (you just use the app, e.g. M365/Teams/Outlook).
- Service model = WHAT is managed. Deployment model = WHERE infra lives. These are independent.
- Office 365 = productivity apps only. Microsoft 365 = Office 365 + Intune (device mgmt) + security + Windows license (E3/E5).
- Cloud Benefits: High Availability (SLA uptime), Scalability (scale up/out), Elasticity (auto-scale up AND down), Reliability (fault tolerance), Agility (provision in minutes), CapEx to OpEx shift.
- Migration "5 Rs": Rehost (lift-and-shift, IaaS, no code changes), Refactor (PaaS, minor changes), Repurchase (replace with SaaS like M365), Retire (decommission), Retain (keep on-prem).
- Exam tip: "minimal changes"/"no re-architecture" = IaaS rehost. "replace on-prem email" = Repurchase (SaaS).

DOMAIN 2 — M365 APPS & SERVICES (45-50% of exam):
- Core apps: Word, Excel, PowerPoint, OneNote, Outlook, Teams.
- Teams = hub for chat, meetings, calling, file sharing. Every Team has a SharePoint site (files) and Exchange mailbox (calendar).
- SharePoint = org/team files & intranet. OneDrive = personal cloud files. Stream = video portal (recordings stored in OneDrive/SharePoint).
- Exchange Online = cloud email & calendaring. Outlook = the client app.
- Microsoft 365 Copilot = AI assistant in M365 apps, uses Microsoft Graph, scoped to user's existing permissions.
- Project management: To Do (personal tasks), Planner (team Kanban boards), Lists (structured data), Loop (real-time collaborative workspaces), Project (complex PM, Gantt charts, add-on).
- Viva suite: Insights (productivity/wellbeing analytics, privacy-protected), Connections (employee intranet in Teams), Learning (training content aggregator), Goals (OKR tracking), Engage (social communities, formerly Yammer).
- Frontline plans: F1 (web/mobile only, lowest cost), F3 (adds desktop apps). Teams Shifts, Walkie Talkie, Task Publishing for frontline.
- Intune = MDM (full device management) + MAM (app-level management for BYOD). Autopatch = automated Windows/Office updates.
- Windows Autopilot = zero-touch device provisioning for new corporate devices.
- Windows 365 = Desktop-as-a-Service (simple, fixed monthly, per-user, SMB). Azure Virtual Desktop (AVD) = VDI (complex, pay-as-you-go, multi-session, enterprise).
- Windows-as-a-Service (WaaS) deployment rings: Insider/Preview → Pilot → Broad.
- M365 Apps update channels: Current (monthly, latest features), Monthly Enterprise (monthly, predictable schedule), Semi-Annual Enterprise (twice/year, max stability).
- M365 Admin Center: manage users, licenses, billing, service health. Adoption Score = M365 usage effectiveness. Secure Score = security posture (different thing!).

DOMAIN 3 — SECURITY & COMPLIANCE (20-25% of exam):
- Zero Trust: "Never trust, always verify." Verify explicitly (MFA+Conditional Access), Least privilege (RBAC+PIM), Assume breach (Defender+Purview).
- Entra ID (formerly Azure AD) = cloud identity service, backbone of every M365 tenant.
- Identity models: Cloud-only, On-premises (AD DS), Hybrid (synced via Entra Cloud Sync).
- Device identity: Entra Registered (BYOD), Entra Joined (cloud-only corporate), Hybrid Entra Joined (both AD+Entra).
- MFA = 2+ verification factors. SSPR = self-service password reset. Conditional Access = policy engine (IF condition THEN action, e.g. require MFA from unknown location).
- Entra External ID: B2B = partners accessing your tenant with their own credentials. B2C = customers signing into your published apps.
- Defender XDR products: Defender for Endpoint (devices, EDR), Defender for Office 365 (email, Safe Links/Attachments), Defender for Identity (on-prem AD threats), Defender for Cloud Apps (CASB, shadow IT).
- CASB four pillars: Visibility, Data Security, Threat Protection, Compliance.
- Secure Score = numeric security posture measurement (0-100+) with actionable recommendations.
- Microsoft Purview: Sensitivity Labels (classify/protect docs), DLP (prevent sensitive data sharing), Insider Risk Management, eDiscovery (legal hold/search), Auditing (Standard=90 days, Premium=1yr+intelligent insights), Retention Policies, Records Management (immutable), Compliance Manager (regulatory compliance scoring).
- Purview = protects organizational data. Priva = protects individuals' personal data (GDPR/CCPA).
- Global Secure Access (GSA) = Security Service Edge (SSE): Entra Internet Access (secure web gateway) + Entra Private Access (VPN-less app access).
- Microsoft Privacy Principles: Control, Transparency, Security, Strong Legal Protections, No Content-Based Targeting, Benefits to You.

DOMAIN 4 — PRICING & LICENSING (10-15% of exam):
- Purchase channels: Enterprise Agreement (EA, 500+ users, 3-year), Cloud Solution Provider (CSP, any size, monthly via partner), MOSP/Web Direct (credit card at microsoft.com).
- Software Assurance (SA) = Volume Licensing only (EA/Open), right to upgrade to latest on-prem software version.
- USL types: Full (no prior licenses), From SA (transitioning SA to cloud), Step Up (upgrade plan tier), Add-On (bolt on capability).
- Business plans: Basic/Standard/Premium, cap at 300 users. Enterprise plans: E3/E5, unlimited users.
- Business Basic = web/mobile apps. Business Standard = adds desktop apps. Business Premium = adds Intune+Defender. E3 = desktop+compliance+Intune. E5 = E3+advanced security+Power BI Pro+Teams Phone.
- Common add-ons: Copilot, Teams Phone, Defender P2, Power BI Pro.
- Support tiers: Basic/Subscription (all plans), Developer, Standard (24/7, 1hr critical), Professional Direct (proactive guidance), Unified/Premier (TAM, enterprise).
- Support requests filed in M365 Admin Center. CSP customers contact partner first.
- FastTrack = free onboarding/adoption service (150+ seats). Not break-fix support.
- SLA: 99.9% uptime standard. Breach → service credits (billing credit, not cash refund).
- Service lifecycle: Private Preview (invite-only, no SLA) → Public Preview (opt-in, no SLA) → GA (production, SLA applies).
- M365 Roadmap portal = track upcoming features. Service Trust Portal = compliance certifications/audit reports.
- Service Health Dashboard = real-time health status. Message Center = planned changes/new features.
- Adoption Score (Admin Center) vs Secure Score (Defender Portal) vs Compliance Manager (Purview) — three different scores measuring different things.`;

  // ─── CONVERSATION STATE ──────────────────────────────────────────────────
  let messages = []; // chat message history: [{role, content}, ...]
  let isOpen = false;
  let engine = null;      // WebLLM engine instance
  let modelReady = false; // true once model is loaded
  let modelLoading = false;

  // ─── DOM CREATION ────────────────────────────────────────────────────────
  function createWidget() {
    // Load CSS
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = 'ms900-chatbot.css';
    document.head.appendChild(link);

    // Toggle button
    const toggle = document.createElement('button');
    toggle.className = 'chatbot-toggle';
    toggle.id = 'chatbot-toggle';
    toggle.innerHTML = '💬';
    toggle.title = 'MS-900 Study Assistant';
    toggle.setAttribute('aria-label', 'Open AI study assistant');
    document.body.appendChild(toggle);

    // Chat panel
    const panel = document.createElement('div');
    panel.className = 'chatbot-panel';
    panel.id = 'chatbot-panel';
    panel.innerHTML = `
      <div class="chatbot-header">
        <div class="chatbot-header-icon">🤖</div>
        <div class="chatbot-header-text">
          <h3>MS-900 Study Assistant</h3>
          <span>Offline AI &middot; Runs in your browser</span>
        </div>
        <div class="chatbot-header-actions">
          <button class="chatbot-header-btn" id="chatbot-clear" title="Clear chat">🗑</button>
          <button class="chatbot-header-btn" id="chatbot-close" title="Close">&times;</button>
        </div>
      </div>
      <div id="chatbot-setup" class="chatbot-setup">
        <div class="chatbot-load-icon">🧠</div>
        <h4>Loading AI Model</h4>
        <p>The AI runs entirely in your browser — no internet needed after the first download (~800 MB, cached automatically).</p>
        <progress id="chatbot-load-progress" value="0" max="1" style="width:100%;margin:10px 0;accent-color:#6366f1"></progress>
        <div id="chatbot-load-text" class="chatbot-key-note" style="text-align:center">Waiting to start...</div>
        <div id="chatbot-load-error" style="display:none;color:#f87171;margin-top:8px;font-size:0.85em;text-align:center"></div>
        <button id="chatbot-load-retry" style="display:none;margin-top:10px">Retry</button>
      </div>
      <div class="chatbot-messages" id="chatbot-messages" style="display:none">
      </div>
      <div class="chatbot-input-area" id="chatbot-input-area" style="display:none">
        <div class="chatbot-input-row">
          <textarea class="chatbot-input" id="chatbot-input" placeholder="Ask about MS-900 topics..." rows="1"></textarea>
          <button class="chatbot-send" id="chatbot-send" title="Send">➤</button>
        </div>
      </div>
    `;
    document.body.appendChild(panel);

    // Wire events
    toggle.addEventListener('click', togglePanel);
    document.getElementById('chatbot-close').addEventListener('click', togglePanel);
    document.getElementById('chatbot-clear').addEventListener('click', clearChat);
    document.getElementById('chatbot-load-retry').addEventListener('click', startModelLoad);
    document.getElementById('chatbot-send').addEventListener('click', sendMessage);

    const input = document.getElementById('chatbot-input');
    input.addEventListener('keydown', function(e) {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        sendMessage();
      }
    });
    // Auto-resize textarea
    input.addEventListener('input', function() {
      this.style.height = 'auto';
      this.style.height = Math.min(this.scrollHeight, 100) + 'px';
    });

    // Start loading the model in the background immediately
    startModelLoad();
  }

  // ─── PANEL TOGGLE ────────────────────────────────────────────────────────
  function togglePanel() {
    isOpen = !isOpen;
    const panel = document.getElementById('chatbot-panel');
    const toggle = document.getElementById('chatbot-toggle');
    panel.classList.toggle('visible', isOpen);
    toggle.classList.toggle('open', isOpen);
    toggle.innerHTML = isOpen ? '✕' : '💬';
    if (isOpen && modelReady) {
      setTimeout(function() { document.getElementById('chatbot-input').focus(); }, 200);
    }
  }

  // ─── MODEL LOADING ───────────────────────────────────────────────────────
  function startModelLoad() {
    if (modelLoading || modelReady) return;
    modelLoading = true;

    // Hide retry button and error if visible
    var retryBtn = document.getElementById('chatbot-load-retry');
    var errEl = document.getElementById('chatbot-load-error');
    if (retryBtn) retryBtn.style.display = 'none';
    if (errEl) errEl.style.display = 'none';

    setLoadText('Loading AI model... (first load may take a while)');

    // Detect GPU f16 compute shader support; fall back to f32 on mobile/limited GPUs
    var modelIdPromise = (navigator.gpu
      ? navigator.gpu.requestAdapter().then(function(adapter) {
          if (!adapter) return MODEL_ID_F32;
          return adapter.features.has('shader-f16') ? MODEL_ID_F16 : MODEL_ID_F32;
        }).catch(function() { return MODEL_ID_F32; })
      : Promise.resolve(MODEL_ID_F32)
    );

    modelIdPromise.then(function(modelId) {
      return import('https://esm.run/@mlc-ai/web-llm').then(function(webllm) {
        var progressCb = function(report) {
          var prog = document.getElementById('chatbot-load-progress');
          var txt = document.getElementById('chatbot-load-text');
          if (prog) prog.value = report.progress || 0;
          if (txt) txt.textContent = report.text || 'Loading...';
        };

        return webllm.CreateMLCEngine(modelId, { initProgressCallback: progressCb });
      });
    }).then(function(eng) {
      engine = eng;
      modelReady = true;
      modelLoading = false;
      showChatView();
      loadHistory();
      addSystemMessage('Ready! Ask me anything about MS-900 exam topics.');
    }).catch(function(err) {
      modelLoading = false;
      var msg = err && err.message ? err.message : String(err);
      var errEl = document.getElementById('chatbot-load-error');
      var retryBtn = document.getElementById('chatbot-load-retry');
      if (errEl) {
        errEl.textContent = 'Failed to load model: ' + msg;
        errEl.style.display = 'block';
      }
      if (retryBtn) retryBtn.style.display = 'inline-block';
      setLoadText('Model failed to load.');
      console.error('WebLLM load error:', err);
    });
  }

  function setLoadText(text) {
    var el = document.getElementById('chatbot-load-text');
    if (el) el.textContent = text;
  }

  // ─── VIEW SWITCHING ──────────────────────────────────────────────────────
  function showChatView() {
    document.getElementById('chatbot-setup').style.display = 'none';
    document.getElementById('chatbot-messages').style.display = 'flex';
    document.getElementById('chatbot-input-area').style.display = 'block';
  }

  // ─── HISTORY PERSISTENCE (sessionStorage) ────────────────────────────────
  function saveHistory() {
    try {
      sessionStorage.setItem(HISTORY_KEY, JSON.stringify(messages));
    } catch(e) {}
  }

  function loadHistory() {
    try {
      var saved = JSON.parse(sessionStorage.getItem(HISTORY_KEY) || '[]');
      if (saved.length > 0) {
        messages = saved;
        var container = document.getElementById('chatbot-messages');
        container.innerHTML = '';
        messages.forEach(function(msg) {
          if (msg.role === 'user') appendBubble(msg.content, 'user');
          else if (msg.role === 'assistant') appendBubble(msg.content, 'bot');
        });
      }
    } catch(e) { messages = []; }
  }

  // ─── PAGE CONTEXT EXTRACTION ─────────────────────────────────────────────
  function extractPageContext() {
    var title = document.title || '';
    var selectors = [
      '.tldr', '.concept-title', '.concept-desc', '.callout-text',
      '.term-key', '.term-val', '.module-title', '.module-desc',
      '.compare-table', '.hero-sub', '.domain-pill',
      '.card-question', '.card-answer', // flashcards
      '.glossary-term', '.glossary-def', '.g-term', '.g-def', // glossary
      'h1', 'h2', 'h3'
    ];

    var contextParts = ['Current page: ' + title];

    selectors.forEach(function(sel) {
      var els = document.querySelectorAll(sel);
      els.forEach(function(el) {
        var text = (el.textContent || '').trim();
        if (text.length > 3 && text.length < 2000) {
          contextParts.push(text);
        }
      });
    });

    document.querySelectorAll('.compare-table').forEach(function(table) {
      var text = table.textContent.replace(/\s+/g, ' ').trim();
      if (text.length > 5) contextParts.push('Table: ' + text);
    });

    var context = contextParts.join('\n');
    if (context.length > 4000) context = context.substring(0, 4000) + '...';
    return context;
  }

  // ─── MESSAGE DISPLAY ────────────────────────────────────────────────────
  function appendBubble(text, type) {
    var container = document.getElementById('chatbot-messages');
    var div = document.createElement('div');
    div.className = 'chatbot-msg ' + type;

    if (type === 'bot') {
      div.innerHTML = formatResponse(text);
    } else {
      div.textContent = text;
    }
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
    return div;
  }

  function addSystemMessage(text) {
    var container = document.getElementById('chatbot-messages');
    var div = document.createElement('div');
    div.className = 'chatbot-msg system';
    div.textContent = text;
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
  }

  function showTyping() {
    var container = document.getElementById('chatbot-messages');
    var div = document.createElement('div');
    div.className = 'chatbot-typing show';
    div.id = 'chatbot-typing';
    div.innerHTML = '<span></span><span></span><span></span>';
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
  }

  function hideTyping() {
    var el = document.getElementById('chatbot-typing');
    if (el) el.remove();
  }

  function formatResponse(text) {
    return text
      .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
      .replace(/\*(.+?)\*/g, '<em>$1</em>')
      .replace(/`(.+?)`/g, '<code>$1</code>')
      .replace(/\n- /g, '\n• ')
      .replace(/\n/g, '<br>');
  }

  // ─── SEND MESSAGE ────────────────────────────────────────────────────────
  function sendMessage() {
    if (!modelReady) return;

    var input = document.getElementById('chatbot-input');
    var text = (input.value || '').trim();
    if (!text) return;

    appendBubble(text, 'user');
    input.value = '';
    input.style.height = 'auto';

    messages.push({ role: 'user', content: text });

    input.disabled = true;
    document.getElementById('chatbot-send').disabled = true;
    showTyping();

    var pageContext = extractPageContext();
    var systemContent = KNOWLEDGE_BASE + '\n\nADDITIONAL CONTEXT FROM CURRENT PAGE:\n' + pageContext;
    var recentMessages = messages.slice(-MAX_HISTORY * 2);

    engine.chat.completions.create({
      messages: [
        { role: 'system', content: systemContent }
      ].concat(recentMessages),
      temperature: 0.3,
      max_tokens: 600
    }).then(function(response) {
      hideTyping();
      var reply = response.choices[0].message.content;
      messages.push({ role: 'assistant', content: reply });
      appendBubble(reply, 'bot');
      saveHistory();
    }).catch(function(err) {
      hideTyping();
      var msg = err && err.message ? err.message : 'Unknown error';
      addSystemMessage('Error: ' + msg);
    }).finally(function() {
      input.disabled = false;
      document.getElementById('chatbot-send').disabled = false;
      input.focus();
    });
  }

  // ─── CLEAR CHAT ──────────────────────────────────────────────────────────
  function clearChat() {
    if (messages.length === 0) return;
    messages = [];
    try { sessionStorage.removeItem(HISTORY_KEY); } catch(e) {}
    document.getElementById('chatbot-messages').innerHTML = '';
    addSystemMessage('Chat cleared. Ask me anything about MS-900!');
  }

  // ─── INIT ────────────────────────────────────────────────────────────────
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', createWidget);
  } else {
    createWidget();
  }

})();
