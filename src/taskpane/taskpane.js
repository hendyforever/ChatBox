(function() {
    // Configuration
    const config = {
        apiKey: localStorage.getItem('deepseek_api_key') || '',
        apiUrl: localStorage.getItem('deepseek_api_url') || 'https://api.deepseek.com/v1',
        model: localStorage.getItem('deepseek_model') || 'deepseek-chat',
        temperature: parseFloat(localStorage.getItem('deepseek_temperature') || '0.7'),
        maxTokens: parseInt(localStorage.getItem('deepseek_max_tokens') || '2048'),
        includeContextByDefault: localStorage.getItem('deepseek_include_context') === 'true',
    };

    // State
    let currentSession = createNewSession();
    let sessions = JSON.parse(localStorage.getItem('chat_sessions') || '[]');
    let isGenerating = false;
    let abortController = null;

    // DOM Elements
    const messageForm = document.getElementById('messageForm');
    const messageInput = document.getElementById('messageInput');
    const messagesContainer = document.getElementById('messagesContainer');
    const historyList = document.getElementById('historyList');
    const contextInfo = document.getElementById('contextInfo');
    const documentContext = document.getElementById('documentContext');
    const includeContextBtn = document.getElementById('includeContextBtn');
    const stopGenerationBtn = document.getElementById('stopGenerationBtn');
    const newChatBtn = document.getElementById('newChatBtn');
    const settingsBtn = document.getElementById('settingsBtn');
    const activeChatTitle = document.getElementById('activeChatTitle');
    function scrollToBottom() {
        messagesContainer.scrollTop = messagesContainer.scrollHeight;
        }
    // Initialize the application
    Office.onReady(async function(info) {
        if (info.host === Office.HostType.Excel) {
            initializeUI();
            updateContextInfo();
            loadSessions();
            
            // Set up event listeners
            messageForm.addEventListener('submit', handleMessageSubmit);
            includeContextBtn.addEventListener('click', toggleIncludeContext);
            stopGenerationBtn.addEventListener('click', stopGeneration);
            newChatBtn.addEventListener('click', createNewChat);
            settingsBtn.addEventListener('click', openSettings);
            
            // Set up context detection
            Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, updateContextInfo);
            Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, updateContextInfo);
        }
    });

    // Initialize UI components
    function initializeUI() {
        // Configure markdown rendering
        marked.setOptions({
            highlight: function(code, lang) {
                if (lang && hljs.getLanguage(lang)) {
                    return hljs.highlight(lang, code).value;
                }
                return hljs.highlightAuto(code).value;
            },
            langPrefix: 'hljs language-',
        });
        
        // Update UI based on config
        includeContextBtn.classList.toggle('active', config.includeContextByDefault);
    }

    // Create a new chat session
    function createNewSession() {
        return {
            id: generateId(),
            title: 'New Chat',
            messages: [],
            createdAt: Date.now(),
            updatedAt: Date.now(),
            context: null,
        };
    }

    // Get current Excel context
    async function getCurrentContext() {
        try {
            return await Excel.run(async function(context) {
                const workbook = context.workbook;
                const worksheet = workbook.getActiveWorksheet();
                const range = context.workbook.getSelectedRange();
                
                range.load('address');
                worksheet.load('name');
                workbook.load('name');
                
                await context.sync();
                
                return {
                    workbookName: workbook.name,
                    worksheetName: worksheet.name,
                    selection: range.address,
                };
            });
        } catch (error) {
            console.error('Error getting context:', error);
            return null;
        }
    }

    // Update context info in UI
    async function updateContextInfo() {
        const context = await getCurrentContext();
        currentSession.context = context;
        
        let contextText = '';
        if (context && context.workbookName) {
            contextText += `Workbook: ${context.workbookName}`;
        }
        if (context && context.worksheetName) {
            contextText += contextText ? ` | Sheet: ${context.worksheetName}` : `Sheet: ${context.worksheetName}`;
        }
        if (context && context.selection) {
            contextText += contextText ? ` | Selection: ${context.selection}` : `Selection: ${context.selection}`;
        }
        
        contextInfo.textContent = contextText || 'No context detected';
        documentContext.textContent = context && context.selection 
            ? `Current selection: ${context.worksheetName}!${context.selection}` 
            : 'No selection detected';
    }

    // Toggle include context setting
    function toggleIncludeContext() {
        config.includeContextByDefault = !config.includeContextByDefault;
        localStorage.setItem('deepseek_include_context', config.includeContextByDefault.toString());
        includeContextBtn.classList.toggle('active', config.includeContextByDefault);
    }

    // Handle message submission
    async function handleMessageSubmit(e) {
        e.preventDefault();
        
        const messageText = messageInput.value.trim();
        if (!messageText || isGenerating) return;
        
        // Add user message
        const userMessage = {
            id: generateId(),
            role: 'user',
            content: messageText,
            timestamp: Date.now(),
        };
        
        addMessageToCurrentSession(userMessage);
        renderMessages();
        messageInput.value = '';
        
        // Generate assistant response
        await generateAssistantResponse();
    }

    // Generate assistant response
    async function generateAssistantResponse() {
        if (isGenerating) return;
        
        isGenerating = true;
        stopGenerationBtn.disabled = false;
        
        try {
            abortController = new AbortController();
            
            // Prepare the messages for the API
            let messages = prepareMessagesForApi();
            
            // Create assistant message placeholder
            const assistantMessageId = generateId();
            const assistantMessage = {
                id: assistantMessageId,
                role: 'assistant',
                content: '',
                timestamp: Date.now(),
            };
            
            addMessageToCurrentSession(assistantMessage);
            renderMessages();
            
            // Scroll to bottom
            scrollToBottom();
            
            // Call DeepSeek API with streaming
            const response = await fetch(`${config.apiUrl}/chat/completions`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${config.apiKey}`,
                },
                body: JSON.stringify({
                    model: config.model,
                    messages: messages,
                    temperature: config.temperature,
                    max_tokens: config.maxTokens,
                    stream: true,
                }),
                signal: abortController.signal,
            });
            
            if (!response.ok) {
                throw new Error(`API request failed with status ${response.status}`);
            }
            
            // Process the streamed response
            const reader = response.body.getReader();
            const decoder = new TextDecoder();
            let partialMessage = '';

            while (true) {
                const { done, value } = await reader.read();
                if (done) break;
                
                const chunk = decoder.decode(value);
                const lines = chunk.split('\n').filter(line => line.trim() !== '');
                
                for (const line of lines) {
                    if (line.startsWith('data:') && !line.includes('[DONE]')) {
                        try {
                            const data = JSON.parse(line.replace('data: ', ''));
                            if (data.choices && data.choices[0].delta.content) {
                                partialMessage += data.choices[0].delta.content;
                                
                                // Update the assistant message with the new content
                                updateMessageContent(assistantMessageId, partialMessage);
                                renderMessages(false);
                                scrollToBottom();
                            }
                        } catch (e) {
                            console.error('Error parsing stream data:', e);
                        }
                    }
                }
            }
            
            // Update the session title if it's the first assistant message
            if (currentSession.messages.filter(m => m.role === 'assistant').length === 1) {
                await generateSessionTitle();
            }
            
            } catch (error) {
                if (error.name === 'AbortError') {
                    console.log('Generation stopped by user');
                } else {
                    console.error('Error generating response:', error);
                    showError('Failed to generate response. Please check your API settings.');
                }
            } finally {
                isGenerating = false;
                stopGenerationBtn.disabled = true;
                abortController = null;
                saveSessions();
            }
            }
            
            // Prepare messages for API including context if needed
            function prepareMessagesForApi() {
            let messages = currentSession.messages.map(m => ({
                role: m.role,
                content: m.role === 'user' && config.includeContextByDefault && currentSession.context
                    ? `${m.content}\n\nCurrent Excel context:\nWorkbook: ${currentSession.context.workbookName}\nWorksheet: ${currentSession.context.worksheetName}\nSelection: ${currentSession.context.selection}`
                    : m.content
            }));
            
            // Ensure we don't exceed token limits by limiting the message history
            const MAX_HISTORY_MESSAGES = 10;
            if (messages.length > MAX_HISTORY_MESSAGES) {
                messages = [
                    ...messages.slice(0, 1), // Keep the first message (system message if exists)
                    ...messages.slice(messages.length - MAX_HISTORY_MESSAGES + 1)
                ];
            }
            
            return messages;
            }
            
            // Stop the generation process
            function stopGeneration() {
            if (abortController) {
                abortController.abort();
                isGenerating = false;
                stopGenerationBtn.disabled = true;
            }
            }
            
            // Generate a title for the session based on the first user message
            async function generateSessionTitle() {
            if (!currentSession.messages.length) return;
            
            const firstMessage = currentSession.messages[0].content;
            const truncatedMessage = firstMessage.length > 50 
                ? firstMessage.substring(0, 50) + '...' 
                : firstMessage;
            
            currentSession.title = truncatedMessage;
            activeChatTitle.textContent = currentSession.title;
            saveSessions();
            renderHistory();
            }
            
            // Add a message to the current session
            function addMessageToCurrentSession(message) {
            currentSession.messages.push(message);
            currentSession.updatedAt = Date.now();
            }
            
            // Update message content
            function updateMessageContent(messageId, newContent) {
            const message = currentSession.messages.find(m => m.id === messageId);
            if (message) {
                message.content = newContent;
                currentSession.updatedAt = Date.now();
            }
            }
            
            // Render all messages in the current session
            function renderMessages(scrollToBottom = true) {
            messagesContainer.innerHTML = '';
            
            for (const message of currentSession.messages) {
                const messageElement = createMessageElement(message);
                messagesContainer.appendChild(messageElement);
            }
            
            if (scrollToBottom) {
                scrollToBottom();
            }
            }
            
            // Create a message element
            function createMessageElement(message) {
            const messageDiv = document.createElement('div');
            messageDiv.className = `message ${message.role}`;
            messageDiv.dataset.messageId = message.id;
            
            const header = document.createElement('div');
            header.className = 'message-header';
            header.innerHTML = `
                <span class="role-badge ${message.role}">${message.role === 'user' ? 'You' : 'DeepSeek'}</span>
                <span class="timestamp">${formatTimestamp(message.timestamp)}</span>
            `;
            
            const contentDiv = document.createElement('div');
            contentDiv.className = 'message-content';
            
            if (message.role === 'assistant' && message.content.includes('```html')) {
                // Special handling for HTML content
                const processedContent = processHtmlContent(message.content);
                contentDiv.innerHTML = processedContent;
            } else {
                // Regular markdown content
                contentDiv.innerHTML = marked(message.content);
            }
            
            // Add event listeners for copy buttons
            const copyButtons = contentDiv.querySelectorAll('.copy-btn');
            copyButtons.forEach(btn => {
                btn.addEventListener('click', () => {
                    const code = btn.parentElement.querySelector('code')?.textContent;
                    if (code) {
                        navigator.clipboard.writeText(code);
                        btn.textContent = 'Copied!';
                        setTimeout(() => btn.textContent = 'Copy', 2000);
                    }
                });
            });
            
            // Add event listeners for run buttons
            const runButtons = contentDiv.querySelectorAll('.run-html-btn');
            runButtons.forEach(btn => {
                btn.addEventListener('click', () => {
                    const html = btn.dataset.html;
                    if (html) {
                        const previewFrame = document.getElementById('htmlPreviewFrame');
                        const previewModal = new bootstrap.Modal(document.getElementById('htmlPreviewModal'));
                        
                        previewFrame.srcdoc = html;
                        previewModal.show();
                    }
                });
            });
            
            messageDiv.appendChild(header);
            messageDiv.appendChild(contentDiv);
            
            return messageDiv;
            }
            
            // Process HTML content to add run buttons
            function processHtmlContent(content) {
            const htmlWithButtons = content.replace(/```html\n([\s\S]*?)\n```/g, function(match, html) {
                const escapedHtml = escapeHtml(html);
                return `
                    <div class="code-block-container">
                        <div class="code-block-header">
                            <button class="btn btn-sm copy-btn">Copy</button>
                            <button class="btn btn-sm btn-success run-html-btn" data-html="${escapedHtml}">Run</button>
                        </div>
                        <pre><code class="language-html">${html}</code></pre>
                    </div>
                `;
            });
            
            return marked(htmlWithButtons);
            }
            
            // Load chat sessions from localStorage
            function loadSessions() {
            const savedSessions = localStorage.getItem('chat_sessions');
            if (savedSessions) {
                sessions = JSON.parse(savedSessions);
                renderHistory();
            }
            
            // If no current session, create one
            if (!currentSession) {
                currentSession = createNewSession();
            }
            }
            
            // Save sessions to localStorage
            function saveSessions() {
            // Update current session in sessions array
            const existingIndex = sessions.findIndex(s => s.id === currentSession.id);
            if (existingIndex >= 0) {
                sessions[existingIndex] = currentSession;
            } else {
                sessions.push(currentSession);
            }
            
            // Sort by updatedAt (newest first)
            sessions.sort((a, b) => b.updatedAt - a.updatedAt);
            
            // Limit to 20 sessions
            if (sessions.length > 20) {
                sessions = sessions.slice(0, 20);
            }
            
            localStorage.setItem('chat_sessions', JSON.stringify(sessions));
            renderHistory();
            }
            
            // Render chat history
            function renderHistory() {
            historyList.innerHTML = '';
            
            for (const session of sessions) {
                const item = document.createElement('a');
                item.href = '#';
                item.className = `list-group-item list-group-item-action ${session.id === currentSession.id ? 'active' : ''}`;
                item.innerHTML = `
                    <div class="d-flex w-100 justify-content-between">
                        <h6 class="mb-1">${session.title}</h6>
                        <small>${formatTimestamp(session.updatedAt, true)}</small>
                    </div>
                    <small>${session.messages.length} messages</small>
                `;
                
                item.addEventListener('click', function(e) {
                    e.preventDefault();
                    loadSession(session.id);
                });
                
                historyList.appendChild(item);
            }
            }
            
            // Load a session by ID
            function loadSession(sessionId) {
            const session = sessions.find(s => s.id === sessionId);
            if (session) {
                currentSession = session;
                activeChatTitle.textContent = session.title;
                renderMessages();
            }
            }
            
            // Create a new chat
            function createNewChat() {
            if (isGenerating) {
                stopGeneration();
            }
            
            currentSession = createNewSession();
            activeChatTitle.textContent = currentSession.title;
            renderMessages();
            saveSessions();
            }
            
            // Open settings dialog
            function openSettings() {
            window.location.href = 'config.html';
            }
            
            // Helper functions
            function generateId() {
            return Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
            }
            
            function formatTimestamp(timestamp, short = false) {
            const date = new Date(timestamp);
            if (short) {
                return date.toLocaleDateString();
            }
            return date.toLocaleString();
            }
            
            
            
            function escapeHtml(unsafe) {
            return unsafe
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#039;");
            }
            
            function showError(message) {
            const errorDiv = document.createElement('div');
            errorDiv.className = 'alert alert-danger';
            errorDiv.textContent = message;
            messagesContainer.appendChild(errorDiv);
            scrollToBottom();
            
            setTimeout(function() {
                errorDiv.remove();
            }, 5000);
            }

            function formatMessage(text) {
                // 配置 marked 选项
                marked.setOptions({
                    breaks: true,
                    gfm: true,
                    highlight: function(code, language) {
                        if (language && hljs.getLanguage(language)) {
                            try {
                                return hljs.highlight(code, { language }).value;
                            } catch (err) {}
                        }
                        return hljs.highlightAuto(code).value;
                    }
                });

                // 自定义渲染器
                const renderer = new marked.Renderer();
                renderer.code = function(code, language) {
                    const validLang = hljs.getLanguage(language) ? language : 'plaintext';
                    const highlightedCode = hljs.highlight(code, { language: validLang }).value;
                    
                    return `<div class="code-block">
                        <div class="code-header">
                            <span class="code-language">${language || 'plaintext'}</span>
                            <div class="code-actions">
                                <button class="code-action-btn copy" onclick="copyCode(this)">复制</button>
                                ${language === 'html' ? '<button class="code-action-btn run">运行</button>' : ''}
                            </div>
                        </div>
                        <pre><code class="hljs ${language}">${highlightedCode}</code></pre>
                    </div>`;
                };

                // 使用配置好的渲染器来渲染 Markdown
                return marked(text, { renderer: renderer });
            }

            function fetchDeepSeekAPIStreaming({ url, headers, payload, streamingText, onComplete }) {
                let fullResponse = '';
                let buffer = '';

                fetch(url, {
                    method: 'POST',
                    headers,
                    body: JSON.stringify(payload)
                })
                .then(response => {
                    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
                    
                    const reader = response.body.getReader();
                    const decoder = new TextDecoder();

                    function processChunk({ done, value }) {
                        if (done) {
                            // 完成时进行最后一次完整渲染
                            streamingText.innerHTML = formatMessage(fullResponse);
                            onComplete?.();
                            return;
                        }

                        // 解码新的数据块
                        const chunk = decoder.decode(value, { stream: true });
                        buffer += chunk;

                        // 处理完整的数据行
                        const lines = buffer.split('\n');
                        buffer = lines.pop(); // 保留不完整的行

                        for (const line of lines) {
                            if (line.startsWith('data:') && !line.includes('[DONE]')) {
                                try {
                                    const data = JSON.parse(line.substring(5).trim());
                                    if (data.choices?.[0]?.delta?.content) {
                                        const content = data.choices[0].delta.content;
                                        fullResponse += content;
                                        
                                        // 实时渲染更新后的完整响应
                                        streamingText.innerHTML = formatMessage(fullResponse);
                                        scrollToBottom();
                                    }
                                } catch (e) {
                                    console.error('Error parsing SSE data:', e);
                                }
                            }
                        }

                        // 继续读取下一个数据块
                        return reader.read().then(processChunk);
                    }

                    return reader.read().then(processChunk);
                })
                .catch(error => {
                    console.error('Error:', error);
                    streamingText.innerHTML += `<br><br>错误: ${error.message}`;
                });
            }

            function copyCode(button) {
                const preElement = button.closest('.code-block').querySelector('pre');
                const code = preElement.textContent;
                
                navigator.clipboard.writeText(code).then(() => {
                    const originalText = button.textContent;
                    button.textContent = '已复制';
                    button.style.color = '#10a37f';
                    
                    setTimeout(() => {
                        button.textContent = originalText;
                        button.style.color = '';
                    }, 2000);
                }).catch(err => {
                    console.error('复制失败:', err);
                });
            }
            })();