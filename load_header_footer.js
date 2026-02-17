// Function to load header and footer content dynamically
function loadHeaderAndFooter() {
    // Wait for the DOM to be fully loaded
    document.addEventListener('DOMContentLoaded', function () {
        // Get header element
        const headerElement = document.getElementById('header');
        
        // Include the header content directly - no placeholder needed since we're setting it immediately
        headerElement.innerHTML = `
            <div id="subBanner" style="width: 940px; height: 40px; display: flex; overflow: hidden; border-radius: 7px 7px 0 0;">
                <div style="width: 25%; height: 100%; background-color: #0A5F23;"></div>
                <div style="width: 75%; height: 100%; background-color: #707276; display: flex; align-items: center; padding-left: 20px;">
                    <span style="color: white; font-size: 22px; font-weight:500; font-family: Arial, sans-serif;">Bin Method Calculator &mdash; RTU Cooling Energy</span>
                </div>
            </div>
        `;
            
        // Add footer content directly
        const footerElement = document.getElementById('footer_new');
        footerElement.innerHTML = `
            <div id="subBanner" style="width: 940px; height: 18px; display: flex; overflow: hidden; border-radius: 0px 0px 7px 7px;">
                <div style="width: 25%; height: 100%; background-color: #0A5F23;"></div>
                <div style="width: 75%; height: 100%; background-color: #707276; display: flex; align-items: center; padding-left: 20px;">
                    <span style="color: white; font-size: 22px; font-weight:500; font-family: Arial, sans-serif;"></span>
                </div>
            </div>
            <p>&nbsp;</p>        
        `;
    });
}

// Call the function to initialize header and footer loading
loadHeaderAndFooter();