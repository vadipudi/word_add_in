/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// Global variables for tracking
let currentTocItems = [];
let autoTrackInterval = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Set up event handlers
    document.getElementById("run").onclick = run;
    document.getElementById("getToc").onclick = getTableOfContents;
    document.getElementById("getSharePointPath").onclick = getSharePointPath;
    document.getElementById("getCurrentSection").onclick = getCurrentSection;
    
    // Set up auto-tracking checkbox
    document.getElementById("autoTrack").onchange = toggleAutoTracking;
  }
});

// Original demo function
export async function run() {
  return Word.run(async (context) => {
    // Insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // Change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}

// Get Table of Contents
export async function getTableOfContents() {
  return Word.run(async (context) => {
    try {
      // Get all paragraphs in the document
      const paragraphs = context.document.body.paragraphs;
      context.load(paragraphs, "text, styleBuiltIn, outlineLevel");
      
      await context.sync();

      const tocItems = [];
      
      // Process paragraphs to find headings
      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        
        // Check if this paragraph is a heading
        if (isHeading(paragraph)) {
          const level = getHeadingLevel(paragraph.styleBuiltIn, paragraph.outlineLevel);
          const text = paragraph.text ? paragraph.text.trim() : '';
          
          if (text) {
            // Load range information for this paragraph
            context.load(paragraph.range, "start, end");
            
            tocItems.push({
              text: text,
              level: level,
              style: paragraph.styleBuiltIn || 'Unknown',
              paragraph: paragraph,
              index: i
            });
          }
        }
      }

      // Sync to get range information
      await context.sync();
      
      // Extract position data safely
      const processedTocItems = [];
      for (let i = 0; i < tocItems.length; i++) {
        const item = tocItems[i];
        try {
          processedTocItems.push({
            text: item.text,
            level: item.level,
            style: item.style,
            start: item.paragraph.range.start,
            end: item.paragraph.range.end,
            index: i
          });
        } catch (rangeError) {
          console.warn("Could not get range for item:", item.text, rangeError);
          // Add item without range info
          processedTocItems.push({
            text: item.text,
            level: item.level,
            style: item.style,
            start: 0,
            end: 0,
            index: i
          });
        }
      }
      
      // Store TOC items globally
      currentTocItems = processedTocItems;

      displayTableOfContents(currentTocItems);
      
    } catch (error) {
      console.error("Error getting table of contents:", error);
      document.getElementById("toc-content").innerHTML = 
        `<p style="color: red;">Error: ${error.message}</p>`;
      document.getElementById("toc-section").style.display = "block";
    }
  });
}

// Helper function to check if paragraph is a heading
function isHeading(paragraph) {
  try {
    // Check by style name
    if (paragraph.styleBuiltIn) {
      const style = paragraph.styleBuiltIn.toString();
      if (style.includes("Heading") || style === "Title") {
        return true;
      }
    }
    
    // Check by outline level (1-9 are typically headings, 10 is body text)
    if (paragraph.outlineLevel !== undefined && paragraph.outlineLevel < 9) {
      return true;
    }
    
    return false;
  } catch (error) {
    console.warn("Error checking if paragraph is heading:", error);
    return false;
  }
}

// Get SharePoint Path
export async function getSharePointPath() {
  try {
    // Show the SharePoint section
    document.getElementById("sharepoint-section").style.display = "block";
    document.getElementById("sharepoint-status").textContent = "Checking...";
    
    // Try to get document properties
    return Word.run(async (context) => {
      try {
        // Get document properties
        const properties = context.document.properties;
        context.load(properties, "title, subject, author, keywords, comments");
        
        await context.sync();
        
        // Try to get the document URL through Office context
        const documentUrl = await getDocumentUrl();
        
        // Update the display
        updateSharePointDisplay(documentUrl, properties);
        
      } catch (error) {
        console.error("Error getting SharePoint info:", error);
        document.getElementById("sharepoint-path").textContent = "Error retrieving path";
        document.getElementById("sharepoint-status").textContent = `Error: ${error.message}`;
      }
    });
    
  } catch (error) {
    console.error("Error in getSharePointPath:", error);
    document.getElementById("sharepoint-status").textContent = `Error: ${error.message}`;
  }
}

// Helper function to get document URL
async function getDocumentUrl() {
  return new Promise((resolve) => {
    try {
      // Try to get the document URL from Office context
      if (Office.context && Office.context.document && Office.context.document.url) {
        resolve(Office.context.document.url);
      } else {
        // Alternative method using window location
        const url = window.location.href;
        if (url.includes('sharepoint') || url.includes('.sharepoint.com')) {
          resolve(url);
        } else {
          resolve("Document URL not available - may not be in SharePoint");
        }
      }
    } catch (error) {
      resolve(`Error getting URL: ${error.message}`);
    }
  });
}

// Helper function to determine heading level
function getHeadingLevel(styleBuiltIn, outlineLevel) {
  try {
    if (styleBuiltIn) {
      const style = styleBuiltIn.toString();
      if (style === "Title") return 0;
      if (style.includes("Heading")) {
        const match = style.match(/Heading(\d+)/);
        if (match && match[1]) {
          return parseInt(match[1]);
        }
      }
    }
    
    // Fallback to outline level
    if (outlineLevel !== undefined && outlineLevel < 9) {
      return outlineLevel;
    }
    
    // Default fallback
    return 1;
  } catch (error) {
    console.warn("Error determining heading level:", error);
    return 1;
  }
}

// Display table of contents
function displayTableOfContents(tocItems) {
  const tocContainer = document.getElementById("toc-content");
  const tocSection = document.getElementById("toc-section");
  
  if (tocItems.length === 0) {
    tocContainer.innerHTML = "<p>No headings found in this document.</p>";
  } else {
    let html = "<ul style='list-style: none; padding-left: 0;'>";
    
    tocItems.forEach((item, index) => {
      const indent = item.level * 20;
      const levelClass = `level-${item.level}`;
      
      html += `
        <li style="margin-left: ${indent}px; margin-bottom: 8px;">
          <div class="toc-item ${levelClass}" data-index="${index}" style="
            padding: 5px; 
            border-left: 3px solid ${getLevelColor(item.level)}; 
            background: ${index % 2 === 0 ? '#f9f9f9' : '#ffffff'};
            border-radius: 3px;
            cursor: pointer;
            transition: all 0.2s ease;
          ">
            <strong>H${item.level}:</strong> ${escapeHtml(item.text)}
            <br><small style="color: #666;">${item.style || 'Unknown style'}</small>
          </div>
        </li>
      `;
    });
    
    html += "</ul>";
    tocContainer.innerHTML = html;
    
    // Add click handlers to TOC items for navigation
    const tocItemElements = tocContainer.querySelectorAll('.toc-item');
    tocItemElements.forEach((element, index) => {
      element.onclick = () => navigateToSection(index);
    });
  }
  
  tocSection.style.display = "block";
}

// Navigate to a specific section
async function navigateToSection(index) {
  if (index < 0 || index >= currentTocItems.length) {
    console.warn("Invalid section index:", index);
    return;
  }
  
  return Word.run(async (context) => {
    try {
      const targetItem = currentTocItems[index];
      
      // Only navigate if we have valid position data
      if (targetItem.start !== undefined && targetItem.start >= 0) {
        // Create a range at the heading position
        const range = context.document.body.getRange();
        range.start = targetItem.start;
        range.end = targetItem.start + Math.max(1, targetItem.text.length);
        
        // Select the range to navigate to it
        range.select();
        
        await context.sync();
        
        // Update current section display after a short delay
        setTimeout(() => {
          getCurrentSection();
        }, 500);
      } else {
        console.warn("No position data available for section:", targetItem.text);
      }
      
    } catch (error) {
      console.error("Error navigating to section:", error);
    }
  });
}

// Update SharePoint display
function updateSharePointDisplay(documentUrl, properties) {
  const pathElement = document.getElementById("sharepoint-path");
  const nameElement = document.getElementById("file-name");
  const statusElement = document.getElementById("sharepoint-status");
  
  // Parse the URL to extract SharePoint information
  if (documentUrl && documentUrl.includes('sharepoint')) {
    pathElement.textContent = documentUrl;
    
    // Extract filename from URL
    const urlParts = documentUrl.split('/');
    const filename = urlParts[urlParts.length - 1] || properties.title || "Unknown";
    nameElement.textContent = filename;
    
    statusElement.textContent = "SharePoint document detected";
    statusElement.style.color = "green";
  } else {
    pathElement.textContent = documentUrl || "Not a SharePoint document";
    nameElement.textContent = properties.title || "Unknown";
    statusElement.textContent = "Not in SharePoint or unable to detect";
    statusElement.style.color = "orange";
  }
}

// Get Current Section
export async function getCurrentSection() {
  return Word.run(async (context) => {
    try {
      // Show position section
      document.getElementById("position-section").style.display = "block";
      document.getElementById("current-section").textContent = "Detecting...";
      
      // Check if we have TOC data
      if (currentTocItems.length === 0) {
        document.getElementById("current-section").textContent = "Please get Table of Contents first";
        document.getElementById("current-position").textContent = "No TOC data available";
        document.getElementById("cursor-info").textContent = "Click 'Get Table of Contents' first";
        return;
      }
      
      // Get current selection/cursor position
      const selection = context.document.getSelection();
      context.load(selection, "start, end, text");
      
      await context.sync();
      
      const currentPosition = selection.start;
      const selectedText = selection.text || '';
      
      // Find which section the cursor is in
      const currentSection = findCurrentSection(currentPosition);
      
      // Update display
      updateCurrentPositionDisplay(currentSection, currentPosition, selectedText);
      
      // Highlight current section in TOC
      highlightCurrentSectionInTOC(currentSection);
      
    } catch (error) {
      console.error("Error getting current section:", error);
      document.getElementById("current-section").textContent = `Error: ${error.message}`;
      document.getElementById("current-position").textContent = "Error occurred";
      document.getElementById("cursor-info").textContent = "Please try again";
    }
  });
}

// Toggle auto-tracking
function toggleAutoTracking() {
  const checkbox = document.getElementById("autoTrack");
  
  if (checkbox.checked) {
    // Start auto-tracking
    autoTrackInterval = setInterval(async () => {
      if (currentTocItems.length > 0) {
        await getCurrentSection();
      }
    }, 2000); // Check every 2 seconds
    
    document.getElementById("position-section").style.display = "block";
  } else {
    // Stop auto-tracking
    if (autoTrackInterval) {
      clearInterval(autoTrackInterval);
      autoTrackInterval = null;
    }
  }
}

// Find which section the current position is in
function findCurrentSection(currentPosition) {
  if (currentTocItems.length === 0) {
    return { text: "No TOC available", level: 0, index: -1 };
  }
  
  // Find the section that contains the current position
  let currentSection = null;
  
  for (let i = 0; i < currentTocItems.length; i++) {
    const item = currentTocItems[i];
    
    // If cursor is after this heading
    if (currentPosition >= item.start) {
      // Check if there's a next heading
      const nextItem = currentTocItems[i + 1];
      
      // If this is the last heading or cursor is before next heading
      if (!nextItem || currentPosition < nextItem.start) {
        currentSection = {
          text: item.text,
          level: item.level,
          style: item.style,
          index: i,
          start: item.start,
          end: item.end
        };
        break;
      }
    }
  }
  
  return currentSection || { text: "Before first heading", level: 0, index: -1 };
}

// Update current position display
function updateCurrentPositionDisplay(section, position, selectedText) {
  document.getElementById("current-section").textContent = section.text || "Unknown";
  document.getElementById("current-position").textContent = `Position: ${position}`;
  
  let cursorInfo = `Position ${position}`;
  if (selectedText && selectedText.trim()) {
    cursorInfo += ` (Selected: "${selectedText.substring(0, 50)}${selectedText.length > 50 ? '...' : ''}")`;
  }
  document.getElementById("cursor-info").textContent = cursorInfo;
}

// Highlight current section in TOC
function highlightCurrentSectionInTOC(currentSection) {
  // Remove previous highlights
  const tocItems = document.querySelectorAll('.toc-item');
  tocItems.forEach(item => {
    item.classList.remove('current-section');
    item.style.backgroundColor = '';
    item.style.fontWeight = '';
  });
  
  // Highlight current section
  if (currentSection && currentSection.index >= 0) {
    const currentItem = tocItems[currentSection.index];
    if (currentItem) {
      currentItem.classList.add('current-section');
      currentItem.style.backgroundColor = '#fff4ce';
      currentItem.style.fontWeight = 'bold';
      currentItem.style.border = '2px solid #ffb900';
      
      // Scroll to current item in TOC
      currentItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  }
}

// Helper functions
function getLevelColor(level) {
  const colors = ['#0078d4', '#107c10', '#d83b01', '#b146c2', '#00bcf2', '#8764b8'];
  return colors[level % colors.length];
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}
