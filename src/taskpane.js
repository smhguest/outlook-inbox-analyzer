/* global Office */

// Initialize when Office.js is ready
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('analyzeBtn').addEventListener('click', analyzeInbox);
        console.log('Inbox Analyzer ready');
    }
});

// Main analysis function
async function analyzeInbox() {
    const statusEl = document.getElementById('status');
    const resultsEl = document.getElementById('results');
    const analyzeBtn = document.getElementById('analyzeBtn');
    
    const folder = document.getElementById('folder').value;
    const limit = parseInt(document.getElementById('limit').value);
    
    // Show loading state
    analyzeBtn.disabled = true;
    statusEl.className = 'status visible';
    statusEl.textContent = 'Scanning emails...';
    resultsEl.innerHTML = '<div class="loading"><div class="spinner"></div><p>Analyzing your mailbox...</p></div>';
    
    try {
        const emails = await fetchEmails(folder, limit);
        
        if (emails.length === 0) {
            resultsEl.innerHTML = '<div class="no-results"><p>No emails found in this folder.</p></div>';
            statusEl.className = 'status';
            analyzeBtn.disabled = false;
            return;
        }
        
        statusEl.textContent = `Analyzing ${emails.length} emails...`;
        
        const analysis = analyzeEmails(emails);
        displayResults(analysis, emails.length);
        
        statusEl.className = 'status';
    } catch (error) {
        console.error('Error analyzing inbox:', error);
        statusEl.className = 'status visible error';
        statusEl.textContent = `Error: ${error.message}`;
        resultsEl.innerHTML = '';
    } finally {
        analyzeBtn.disabled = false;
    }
}

// Fetch emails from the selected folder using EWS
async function fetchEmails(folderName, limit) {
    return new Promise((resolve, reject) => {
        const mailbox = Office.context.mailbox;
        
        // Build EWS request to find items
        const ewsRequest = `<?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                           xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                           xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
                <soap:Header>
                    <t:RequestServerVersion Version="Exchange2013"/>
                </soap:Header>
                <soap:Body>
                    <m:FindItem Traversal="Shallow">
                        <m:ItemShape>
                            <t:BaseShape>Default</t:BaseShape>
                            <t:AdditionalProperties>
                                <t:FieldURI FieldURI="message:From"/>
                                <t:FieldURI FieldURI="message:Sender"/>
                                <t:FieldURI FieldURI="item:DateTimeReceived"/>
                                <t:FieldURI FieldURI="item:Subject"/>
                            </t:AdditionalProperties>
                        </m:ItemShape>
                        <m:IndexedPageItemView MaxEntriesReturned="${limit}" Offset="0" BasePoint="Beginning"/>
                        <m:SortOrder>
                            <t:FieldOrder Order="Descending">
                                <t:FieldURI FieldURI="item:DateTimeReceived"/>
                            </t:FieldOrder>
                        </m:SortOrder>
                        <m:ParentFolderIds>
                            <t:DistinguishedFolderId Id="${folderName}"/>
                        </m:ParentFolderIds>
                    </m:FindItem>
                </soap:Body>
            </soap:Envelope>`;
        
        mailbox.makeEwsRequestAsync(ewsRequest, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const emails = parseEwsResponse(result.value);
                resolve(emails);
            } else {
                // Fallback: try REST API if EWS fails
                fetchEmailsREST(limit)
                    .then(resolve)
                    .catch(() => reject(new Error(result.error.message || 'Failed to fetch emails')));
            }
        });
    });
}

// Parse EWS XML response
function parseEwsResponse(xmlString) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlString, 'text/xml');
    const emails = [];
    
    // Find all Message elements
    const messages = xmlDoc.getElementsByTagNameNS(
        'http://schemas.microsoft.com/exchange/services/2006/types',
        'Message'
    );
    
    for (let i = 0; i < messages.length; i++) {
        const msg = messages[i];
        
        // Get sender info
        const fromEl = msg.getElementsByTagNameNS(
            'http://schemas.microsoft.com/exchange/services/2006/types',
            'From'
        )[0];
        
        let senderName = 'Unknown';
        let senderEmail = 'unknown@unknown.com';
        
        if (fromEl) {
            const mailbox = fromEl.getElementsByTagNameNS(
                'http://schemas.microsoft.com/exchange/services/2006/types',
                'Mailbox'
            )[0];
            
            if (mailbox) {
                const nameEl = mailbox.getElementsByTagNameNS(
                    'http://schemas.microsoft.com/exchange/services/2006/types',
                    'Name'
                )[0];
                const emailEl = mailbox.getElementsByTagNameNS(
                    'http://schemas.microsoft.com/exchange/services/2006/types',
                    'EmailAddress'
                )[0];
                
                if (nameEl) senderName = nameEl.textContent;
                if (emailEl) senderEmail = emailEl.textContent.toLowerCase();
            }
        }
        
        // Get subject and date
        const subjectEl = msg.getElementsByTagNameNS(
            'http://schemas.microsoft.com/exchange/services/2006/types',
            'Subject'
        )[0];
        const dateEl = msg.getElementsByTagNameNS(
            'http://schemas.microsoft.com/exchange/services/2006/types',
            'DateTimeReceived'
        )[0];
        
        emails.push({
            senderName: senderName,
            senderEmail: senderEmail,
            subject: subjectEl ? subjectEl.textContent : '(No subject)',
            date: dateEl ? new Date(dateEl.textContent) : new Date()
        });
    }
    
    return emails;
}

// Fallback REST API method (for newer Outlook versions)
async function fetchEmailsREST(limit) {
    return new Promise((resolve, reject) => {
        // This uses the Office.js REST API wrapper
        if (Office.context.mailbox.restUrl) {
            const restUrl = Office.context.mailbox.restUrl;
            
            Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const token = result.value;
                    
                    fetch(`${restUrl}/v2.0/me/messages?$top=${limit}&$select=from,subject,receivedDateTime&$orderby=receivedDateTime desc`, {
                        headers: {
                            'Authorization': `Bearer ${token}`,
                            'Content-Type': 'application/json'
                        }
                    })
                    .then(response => response.json())
                    .then(data => {
                        const emails = data.value.map(msg => ({
                            senderName: msg.from?.emailAddress?.name || 'Unknown',
                            senderEmail: (msg.from?.emailAddress?.address || 'unknown@unknown.com').toLowerCase(),
                            subject: msg.subject || '(No subject)',
                            date: new Date(msg.receivedDateTime)
                        }));
                        resolve(emails);
                    })
                    .catch(reject);
                } else {
                    reject(new Error('Could not get REST token'));
                }
            });
        } else {
            reject(new Error('REST API not available'));
        }
    });
}

// Analyze email data
function analyzeEmails(emails) {
    const senderMap = new Map();
    const domainMap = new Map();
    let oldestDate = new Date();
    let newestDate = new Date(0);
    
    emails.forEach(email => {
        // Track by email address
        const key = email.senderEmail;
        if (senderMap.has(key)) {
            const sender = senderMap.get(key);
            sender.count++;
            if (email.date < sender.firstEmail) sender.firstEmail = email.date;
            if (email.date > sender.lastEmail) sender.lastEmail = email.date;
        } else {
            senderMap.set(key, {
                email: email.senderEmail,
                name: email.senderName,
                count: 1,
                firstEmail: email.date,
                lastEmail: email.date
            });
        }
        
        // Track by domain
        const domain = email.senderEmail.split('@')[1] || 'unknown';
        domainMap.set(domain, (domainMap.get(domain) || 0) + 1);
        
        // Track date range
        if (email.date < oldestDate) oldestDate = email.date;
        if (email.date > newestDate) newestDate = email.date;
    });
    
    // Convert to sorted arrays
    const senders = Array.from(senderMap.values())
        .sort((a, b) => b.count - a.count);
    
    const domains = Array.from(domainMap.entries())
        .map(([domain, count]) => ({ domain, count }))
        .sort((a, b) => b.count - a.count);
    
    return {
        totalEmails: emails.length,
        uniqueSenders: senders.length,
        uniqueDomains: domains.length,
        dateRange: { oldest: oldestDate, newest: newestDate },
        senders: senders,
        domains: domains.slice(0, 10)
    };
}

// Display results in the UI
function displayResults(analysis, totalScanned) {
    const resultsEl = document.getElementById('results');
    const maxCount = analysis.senders[0]?.count || 1;
    
    // Calculate date range string
    const daysDiff = Math.ceil((analysis.dateRange.newest - analysis.dateRange.oldest) / (1000 * 60 * 60 * 24));
    const dateRangeStr = daysDiff === 0 ? 'today' : `${daysDiff} days`;
    
    let html = `
        <div class="summary">
            <h2>📈 Summary</h2>
            <div class="summary-stats">
                <div class="stat">
                    <div class="stat-value">${analysis.totalEmails}</div>
                    <div class="stat-label">Emails Scanned</div>
                </div>
                <div class="stat">
                    <div class="stat-value">${analysis.uniqueSenders}</div>
                    <div class="stat-label">Unique Senders</div>
                </div>
                <div class="stat">
                    <div class="stat-value">${analysis.uniqueDomains}</div>
                    <div class="stat-label">Unique Domains</div>
                </div>
                <div class="stat">
                    <div class="stat-value">${dateRangeStr}</div>
                    <div class="stat-label">Time Span</div>
                </div>
            </div>
        </div>
        
        <div class="sender-list">
            <h2>📬 Top Senders</h2>
            ${analysis.senders.slice(0, 25).map(sender => `
                <div class="sender-item" title="Click to search for emails from ${sender.email}">
                    <div class="sender-info">
                        <div class="sender-name">${escapeHtml(sender.name)}</div>
                        <div class="sender-email">${escapeHtml(sender.email)}</div>
                        <div class="sender-bar">
                            <div class="sender-bar-fill" style="width: ${(sender.count / maxCount) * 100}%"></div>
                        </div>
                    </div>
                    <span class="sender-count">${sender.count}</span>
                </div>
            `).join('')}
        </div>
        
        <button class="export-btn" onclick="exportToCSV()">📥 Export to CSV</button>
    `;
    
    resultsEl.innerHTML = html;
    
    // Store analysis for export
    window.currentAnalysis = analysis;
}

// Export results to CSV
function exportToCSV() {
    if (!window.currentAnalysis) return;
    
    const analysis = window.currentAnalysis;
    let csv = 'Sender Name,Sender Email,Email Count,First Email,Last Email\n';
    
    analysis.senders.forEach(sender => {
        csv += `"${sender.name.replace(/"/g, '""')}","${sender.email}",${sender.count},"${sender.firstEmail.toLocaleDateString()}","${sender.lastEmail.toLocaleDateString()}"\n`;
    });
    
    // Create download link
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `inbox-analysis-${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Helper to escape HTML
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
