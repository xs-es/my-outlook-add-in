// Signature Auto-Loading Manager for Issue #5803
let signatureLoaded = false;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('Signature Manager initialized');
        initializeSignatureLoading();
    }
});

function initializeSignatureLoading() {
    // Multiple triggers to ensure signature loads in preview pane
    setTimeout(() => loadSignatureIfNeeded(), 100);
    setTimeout(() => loadSignatureIfNeeded(), 1000);
    setTimeout(() => loadSignatureIfNeeded(), 3000);
    
    // Listen for item changes
    if (Office.context.mailbox.item) {
        try {
            Office.context.mailbox.item.addHandlerAsync(
                Office.EventType.ItemChanged,
                loadSignatureIfNeeded
            );
        } catch (e) {
            console.log('ItemChanged handler not supported:', e);
        }
    }
    
    // Start polling for preview pane
    if (isPreviewPane()) {
        console.log('Preview pane detected, starting polling');
        startSignaturePolling();
    }
}

function isPreviewPane() {
    try {
        return window.parent !== window && window.frameElement !== null;
    } catch (e) {
        return true; // Assume preview pane if cross-origin
    }
}

function loadSignatureIfNeeded() {
    if (signatureLoaded) return;
    
    try {
        Office.context.mailbox.item.body.getAsync(
            Office.CoercionType.Html,
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const bodyContent = result.value;
                    
                    if (!hasSignature(bodyContent)) {
                        console.log('Adding signature...');
                        addSignature();
                    } else {
                        console.log('Signature already present');
                        signatureLoaded = true;
                    }
                }
            }
        );
    } catch (error) {
        console.log('Error checking for signature:', error);
    }
}

function hasSignature(bodyContent) {
    return bodyContent.includes('<!-- SIGNATURE_MARKER -->') ||
           bodyContent.includes('Best regards,'); // Adjust this to your signature
}

function addSignature() {
    const signatureHtml = getSignatureHtml();
    
    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                let bodyContent = result.value;
                
                // Add signature with marker
                const signatureWithMarker = `
                    <br><br>
                    <!-- SIGNATURE_MARKER -->
                    ${signatureHtml}
                `;
                
                bodyContent += signatureWithMarker;
                
                Office.context.mailbox.item.body.setAsync(
                    bodyContent,
                    { coercionType: Office.CoercionType.Html },
                    (setResult) => {
                        if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log('Signature added successfully');
                            signatureLoaded = true;
                        }
                    }
                );
            }
        }
    );
}

function getSignatureHtml() {
    // REPLACE THIS WITH YOUR ACTUAL SIGNATURE
    return `
        <div style="font-family: Arial, sans-serif; font-size: 14px;">
            <p>Best regards,</p>
            <p><strong>Your Name</strong></p>
            <p>Your Title</p>
            <p>Your Company</p>
            <p>Email: your.email@company.com</p>
        </div>
    `;
}

function startSignaturePolling() {
    let attempts = 0;
    const maxAttempts = 10;
    
    const pollInterval = setInterval(() => {
        attempts++;
        console.log(`Polling attempt ${attempts}`);
        
        if (attempts >= maxAttempts || signatureLoaded) {
            clearInterval(pollInterval);
            return;
        }
        
        loadSignatureIfNeeded();
    }, 2000);
}