(function() {
    // Initialize when Office is ready
    Office.onReady(function() {
        // Load current settings
        loadSettings();
        
        // Set up form submission
        const settingsForm = document.getElementById('settingsForm');
        settingsForm.addEventListener('submit', saveSettings);
        
        // Set up back button
        const backButton = document.getElementById('backButton');
        backButton.addEventListener('click', function() {
            window.location.href = 'chat.html';
        });
    });

    function loadSettings() {
        const apiKeyInput = document.getElementById('apiKey');
        const apiUrlInput = document.getElementById('apiUrl');
        const modelInput = document.getElementById('model');
        const temperatureInput = document.getElementById('temperature');
        const maxTokensInput = document.getElementById('maxTokens');
        const includeContextInput = document.getElementById('includeContext');
        
        apiKeyInput.value = localStorage.getItem('deepseek_api_key') || '';
        apiUrlInput.value = localStorage.getItem('deepseek_api_url') || 'https://api.deepseek.com/v1';
        modelInput.value = localStorage.getItem('deepseek_model') || 'deepseek-chat';
        temperatureInput.value = localStorage.getItem('deepseek_temperature') || '0.7';
        maxTokensInput.value = localStorage.getItem('deepseek_max_tokens') || '2048';
        includeContextInput.checked = localStorage.getItem('deepseek_include_context') === 'true';
    }

    function saveSettings(e) {
        e.preventDefault();
        
        const apiKeyInput = document.getElementById('apiKey');
        const apiUrlInput = document.getElementById('apiUrl');
        const modelInput = document.getElementById('model');
        const temperatureInput = document.getElementById('temperature');
        const maxTokensInput = document.getElementById('maxTokens');
        const includeContextInput = document.getElementById('includeContext');
        
        localStorage.setItem('deepseek_api_key', apiKeyInput.value);
        localStorage.setItem('deepseek_api_url', apiUrlInput.value);
        localStorage.setItem('deepseek_model', modelInput.value);
        localStorage.setItem('deepseek_temperature', temperatureInput.value);
        localStorage.setItem('deepseek_max_tokens', maxTokensInput.value);
        localStorage.setItem('deepseek_include_context', includeContextInput.checked.toString());
        
        alert('Settings saved successfully!');
        window.location.href = 'taskpane.html';
    }
})();