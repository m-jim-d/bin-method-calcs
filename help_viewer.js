/**
 * Modal Help Viewer
 * Provides a modern modal dialog approach for displaying help content
 * instead of using window.open() to create new browser windows.
 */

// Global variables for modal components
var helpModalOverlay = null;
var helpModalFrame = null;
var currentHelpHash = '';
var escapeKeyListenerAdded = false; // Track if escape listener is already added
var _helpViewerBase = (function() {
    var s = document.currentScript;
    if (!s || !s.src) return '';
    return s.src.substring(0, s.src.lastIndexOf('/') + 1);
})();

// Ensure the modal is initialized when the DOM is ready
document.addEventListener('DOMContentLoaded', function() {
    // Initialize modal HTML structure
    initHelpModalHTML();
});

/**
 * Creates the modal HTML structure if it doesn't exist in the document
 */
function initHelpModalHTML() {
    // Check if modal already exists
    if (document.getElementById('helpModalOverlay')) {
        return;
    }
    
    // Create modal elements
    helpModalOverlay = document.createElement('div');
    helpModalOverlay.id = 'helpModalOverlay';
    helpModalOverlay.className = 'modal-overlay';
    helpModalOverlay.onclick = closeHelpModal; // Close when clicking outside content
    
    var modalContent = document.createElement('div');
    modalContent.className = 'modal-content';
    modalContent.onclick = function(event) {
        event.stopPropagation(); // Prevent closing when clicking inside content
    };
    
    var closeButton = document.createElement('button');
    closeButton.type = 'button';
    closeButton.className = 'modal-close-button';
    closeButton.innerHTML = '×';
    closeButton.title = 'Close (Esc)';
    closeButton.onclick = closeHelpModal;
    
    // Create a container for the iframe with padding
    var iframeContainer = document.createElement('div');
    iframeContainer.className = 'iframe-container';
    
    helpModalFrame = document.createElement('iframe');
    helpModalFrame.id = 'helpModalFrame';
    helpModalFrame.src = 'about:blank';
    helpModalFrame.title = 'Help Content';
    helpModalFrame.style.width = '100%';
    helpModalFrame.style.height = '100%';
    helpModalFrame.style.border = 'none';
    
    // Add the iframe to the container
    iframeContainer.appendChild(helpModalFrame);
    
    // Assemble the modal structure
    modalContent.appendChild(closeButton);
    modalContent.appendChild(iframeContainer);
    helpModalOverlay.appendChild(modalContent);
    
    // Add modal to the document body
    document.body.appendChild(helpModalOverlay);
    
    // Add CSS if not already present
    addModalCSS();
    
    // Set up event listener for Escape key (only once)
    if (!escapeKeyListenerAdded) {
        // Add to both document and window for better coverage
        document.addEventListener('keydown', handleHelpModalEscapeKey, true);
        window.addEventListener('keydown', handleHelpModalEscapeKey, true);
        escapeKeyListenerAdded = true;
    }
}

/**
 * Adds the required CSS for the modal if not already present
 */
function addModalCSS() {
    // Check if the CSS is already added
    var cssId = 'modalHelpViewerCSS';
    if (document.getElementById(cssId)) {
        return;
    }
    
    // Create style element
    var head = document.getElementsByTagName('head')[0];
    var style = document.createElement('style');
    style.id = cssId;
    style.type = 'text/css';
    
    // Modal CSS rules
    var css = `
    /* Modal Styles for Help_Controls.html */
    .modal-overlay {
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background-color: rgba(0, 0, 0, 0.6);
      z-index: 10000;
      display: none;
      justify-content: center;
      align-items: center;
      padding: 20px;
      box-sizing: border-box;
    }
    
    .modal-content {
      background-color: #fff;
      padding: 0;
      border-radius: 8px;
      box-shadow: 0 5px 15px rgba(0,0,0,0.3);
      width: 90%;
      max-width: 740px;
      height: 90%;
      max-height: 700px;
      position: relative;
      display: flex;
      flex-direction: column;
      overflow: hidden;
    }
    
    .modal-close-button {
      position: absolute;
      top: 8px;
      right: 8px;
      font-size: 28px;
      font-weight: normal;
      line-height: 1;
      color: #666;
      background: rgba(255, 255, 255, 0.9);
      border: 1px solid #ddd;
      cursor: pointer;
      padding: 0px 3px;
      z-index: 10;
      border-radius: 4px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    .modal-close-button:hover,
    .modal-close-button:focus {
      background: rgba(240, 240, 240, 0.95);
      color: #000;
      border-color: #bbb;
      text-decoration: none;
      outline: none;
    }
    
    .iframe-container {
      flex-grow: 1;
      display: flex;
      flex-direction: column;
      overflow: hidden;
      position: relative;
      padding: 15px 15px 10px 15px; /* Reduced top padding to minimize dead space */
    }
    
    #helpModalFrame {
      flex-grow: 1;
      border: none;
      width: 100%;
      height: 100%;
    }`;
    
    // Add CSS to the style element
    if (style.styleSheet) {
        style.styleSheet.cssText = css; // IE
    } else {
        style.appendChild(document.createTextNode(css));
    }
    
    // Add style to head
    head.appendChild(style);
}

/**
 * Opens the help modal with the specified hashmark
 * @param {string} theHashMark - The hashmark to navigate to in the help document
 */
function openHelpModal(theHashMark) {
    // Ensure modal elements exist
    initHelpModalHTML();
    
    // Store current hash
    currentHelpHash = theHashMark || '';
    
    // Set the iframe source with hashmark and cache-busting parameter
    var helpUrl = _helpViewerBase + "Help_Controls.html";
    helpUrl += "?nocache=1234"; // manually update this number to force a refresh
    
    // Add hashmark if provided (after the cache parameter)
    if (theHashMark && theHashMark !== '') {
        helpUrl += '#' + theHashMark;
    }
    
    // Load content and show modal
    helpModalFrame.src = helpUrl;
    helpModalOverlay.style.display = 'flex';
    document.body.style.overflow = 'hidden'; // Prevent background scrolling
    
    // Focus the modal overlay instead of iframe for better key handling
    setTimeout(function() {
        if (helpModalOverlay) {
            helpModalOverlay.focus();
            // Also add escape key listener to the modal overlay
            helpModalOverlay.addEventListener('keydown', handleHelpModalEscapeKey);
        }
    }, 100);
}

/**
 * Closes the help modal
 */
function closeHelpModal() {
    if (helpModalOverlay) {
        helpModalOverlay.style.display = 'none';
        
        // Clear iframe content to stop any videos/audio
        if (helpModalFrame) {
            helpModalFrame.src = 'about:blank';
        }
        
        // Restore background scrolling
        document.body.style.overflow = '';
    }
}

/**
 * Handles Escape key press to close the modal
 * @param {KeyboardEvent} event - The keyboard event
 */
function handleHelpModalEscapeKey(event) {
    // Check if the modal is open and Escape key is pressed
    if ((event.key === 'Escape' || event.keyCode === 27 || event.which === 27)) {
        if (helpModalOverlay && helpModalOverlay.style.display === 'flex') {
            event.preventDefault(); // Prevent other Escape handlers
            event.stopPropagation(); // Stop event from bubbling
            closeHelpModal();
        }
    }
}

/**
 * Replacement for the legacy OpenWindow function
 * @param {string} theHashMark - The hashmark to navigate to
 */
function OpenWindow(theHashMark) {
    openHelpModal(theHashMark);
}