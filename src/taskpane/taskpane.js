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
    
      // Add event listeners
  document.getElementById("getToc").onclick = getTableOfContents;
  document.getElementById("testMinimal").onclick = testMinimal;
  document.getElementById("getSharePointPath").onclick = getSharePointPath;
  document.getElementById("getCurrentSection").onclick = () => {
    const selectedStyle = document.getElementById("sectionHeadingStyle").value;
    getCurrentSection(selectedStyle);
  };
    
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

// Get Table of Contents - most efficient approach using content controls
export async function getTableOfContents() {
  return Word.run(async (context) => {
    try {
      console.log("Starting getTableOfContents - using content controls approach...");
      
      // Clear previous results
      currentTocItems = [];
      document.getElementById("toc-content").innerHTML = "Loading...";
      document.getElementById("toc-section").style.display = "block";
      
      const tocItems = [];
      
      try {
        // Method 1: Try to get built-in Table of Contents if it exists
        const contentControls = context.document.contentControls;
        context.load(contentControls, "items");
        await context.sync();
        
        console.log(`Found ${contentControls.items.length} content controls`);
        
        // Look for TOC content controls
        let foundTOC = false;
        for (const control of contentControls.items) {
          context.load(control, "type, title, text");
        }
        await context.sync();
        
        for (const control of contentControls.items) {
          if (control.title && control.title.toLowerCase().includes('table') && control.title.toLowerCase().includes('contents')) {
            console.log("Found existing TOC content control!");
            const tocText = control.text || '';
            // Parse the existing TOC (basic approach)
            const lines = tocText.split('\n');
            lines.forEach((line, index) => {
              const trimmed = line.trim();
              if (trimmed && !trimmed.match(/^\d+$/) && trimmed.length > 1) {
                // Estimate level based on indentation or content
                const level = line.length - line.trimStart().length > 10 ? 2 : 1;
                tocItems.push({
                  text: trimmed,
                  level: level,
                  style: `Heading ${level}`,
                  index: tocItems.length
                });
              }
            });
            foundTOC = true;
            break;
          }
        }
        
        if (!foundTOC) {
          console.log("No existing TOC found, scanning document structure...");
          // Fallback: Scan document more efficiently
          await scanForHeadingsEfficiently(context, tocItems);
        }
        
      } catch (contentControlError) {
        console.warn("Content control approach failed, using fallback:", contentControlError);
        await scanForHeadingsEfficiently(context, tocItems);
      }
      
      console.log(`Found ${tocItems.length} headings total`);
      
      // Store TOC data
      currentTocItems = tocItems;
      displayTableOfContents(currentTocItems);
      
    } catch (error) {
      console.error("Error getting table of contents:", error);
      console.error("Error details:", {
        name: error.name,
        message: error.message,
        code: error.code,
        traceMessages: error.traceMessages
      });
      
      document.getElementById("toc-content").innerHTML = 
        `<p style="color: red;">Error: ${error.message}</p>
         <p style="color: red; font-size: 12px;">Details: ${error.name} (${error.code})</p>`;
      document.getElementById("toc-section").style.display = "block";
    }
  });
}

// Helper function to scan for headings more efficiently
async function scanForHeadingsEfficiently(context, tocItems) {
  // Only get paragraphs that are likely to be headings based on outline level
  const body = context.document.body;
  const ranges = body.getRange();
  context.load(ranges, "paragraphs");
  await context.sync();
  
  // Load only first few paragraphs to test the approach
  const sampleSize = Math.min(50, ranges.paragraphs.items.length); // Limit to first 50 paragraphs
  console.log(`Checking first ${sampleSize} paragraphs for headings...`);
  
  for (let i = 0; i < sampleSize; i++) {
    const para = ranges.paragraphs.items[i];
    context.load(para, "text, styleBuiltIn, outlineLevel");
  }
  
  await context.sync();
  
  for (let i = 0; i < sampleSize; i++) {
    const para = ranges.paragraphs.items[i];
    const text = para.text ? para.text.trim() : '';
    const style = para.styleBuiltIn ? para.styleBuiltIn.toString() : '';
    const outlineLevel = para.outlineLevel;
    
    // Only process if it's clearly a heading
    if (text && text.length < 200 && (
      style.includes('Heading') || 
      style === 'Title' || 
      (outlineLevel < 9 && text.length < 100)
    )) {
      const level = style === 'Title' ? 0 : 
                   style.includes('Heading') ? getHeadingLevelSimple(style) : 
                   outlineLevel;
      
      tocItems.push({
        text: text,
        level: level,
        style: style,
        index: tocItems.length
      });
      
      console.log(`Found heading: "${text}" (level ${level}, style: ${style})`);
    }
  }
}

// Minimal test function to debug Word API
export async function testMinimal() {
  return Word.run(async (context) => {
    try {
      console.log("Testing minimal Word API access...");
      
      // Just try to get the document body
      const body = context.document.body;
      context.load(body, "text");
      
      await context.sync();
      
      console.log("Body text length:", body.text.length);
      console.log("First 100 characters:", body.text.substring(0, 100));
      
      document.getElementById("toc-content").innerHTML = 
        `<p style="color: green;">✅ Minimal test passed!</p>
         <p>Document has ${body.text.length} characters</p>`;
      
    } catch (error) {
      console.error("Minimal test failed:", error);
      document.getElementById("toc-content").innerHTML = 
        `<p style="color: red;">❌ Minimal test failed: ${error.message}</p>`;
    }
  });
}

// Simplified heading level function
function getHeadingLevelSimple(style) {
  if (style === 'Title') return 0;
  if (style.includes('Heading')) {
    const match = style.match(/(\d+)/);
    return match ? parseInt(match[1]) : 1;
  }
  return 1;
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

// Get Current Section - finds which section the cursor is in based on heading boundaries
export async function getCurrentSection(sectionHeadingStyle = 'Heading 1') {
  return Word.run(async (context) => {
    try {
      console.log(`Finding current section based on ${sectionHeadingStyle} boundaries...`);
      
      // Show position section
      document.getElementById("position-section").style.display = "block";
      document.getElementById("current-section").textContent = "Detecting...";
      
      // Get cursor position
      const selection = context.document.getSelection();
      context.load(selection, "start");
      await context.sync();
      
      const cursorPosition = selection.start;
      console.log(`Cursor position: ${cursorPosition}`);
      
      // Find all headings of the specified level (section boundaries)
      const sectionHeadings = await findHeadingsByStyle(context, sectionHeadingStyle);
      console.log(`Found ${sectionHeadings.length} section headings`);
      
      if (sectionHeadings.length === 0) {
        document.getElementById("current-section").textContent = `No ${sectionHeadingStyle} found`;
        document.getElementById("current-position").textContent = "Cannot determine section";
        document.getElementById("cursor-info").textContent = `Add some ${sectionHeadingStyle} headings to use this feature`;
        return;
      }
      
      // Find which section the cursor is in
      const currentSection = findSectionFromPosition(cursorPosition, sectionHeadings);
      
      // Update display
      if (currentSection) {
        document.getElementById("current-section").textContent = `Section: ${currentSection.title}`;
        document.getElementById("current-position").textContent = `Position: ${cursorPosition} (in ${currentSection.title})`;
        document.getElementById("cursor-info").textContent = `Section starts at position ${currentSection.start}`;
      } else {
        document.getElementById("current-section").textContent = "Before first section";
        document.getElementById("current-position").textContent = `Position: ${cursorPosition}`;
        document.getElementById("cursor-info").textContent = `Before first ${sectionHeadingStyle}`;
      }
      
      return currentSection;
      
    } catch (error) {
      console.error("Error getting current section:", error);
      document.getElementById("current-section").textContent = `Error: ${error.message}`;
      document.getElementById("current-position").textContent = "Error occurred";
      document.getElementById("cursor-info").textContent = "Please try again";
      return null;
    }
  });
}

// Helper function to find headings by style
async function findHeadingsByStyle(context, headingStyle) {
  const headings = [];
  
  try {
    // Get document body and scan for headings efficiently
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    
    // Load paragraphs metadata first
    context.load(paragraphs, "items");
    await context.sync();
    
    // Limit scan to avoid performance issues
    const maxParagraphs = Math.min(200, paragraphs.items.length);
    console.log(`Scanning first ${maxParagraphs} paragraphs for ${headingStyle}...`);
    
    // Load properties for paragraphs in batches
    for (let i = 0; i < maxParagraphs; i++) {
      context.load(paragraphs.items[i], "text, styleBuiltIn, getRange");
    }
    await context.sync();
    
    // Get ranges for position information
    for (let i = 0; i < maxParagraphs; i++) {
      const paragraph = paragraphs.items[i];
      const style = paragraph.styleBuiltIn ? paragraph.styleBuiltIn.toString() : '';
      
      if (style === headingStyle || 
          (headingStyle === 'Heading 1' && style.includes('Heading1')) ||
          (headingStyle === 'Heading 2' && style.includes('Heading2'))) {
        
        const text = paragraph.text ? paragraph.text.trim() : '';
        if (text) {
          // Get the range to find position
          const range = paragraph.getRange();
          context.load(range, "start, end");
          await context.sync();
          
          headings.push({
            title: text,
            start: range.start,
            end: range.end,
            style: style
          });
          
          console.log(`Found section heading: "${text}" at position ${range.start}`);
        }
      }
    }
    
  } catch (error) {
    console.error("Error finding headings:", error);
  }
  
  // Sort by position
  headings.sort((a, b) => a.start - b.start);
  return headings;
}

// Helper function to find which section contains a given position
function findSectionFromPosition(cursorPosition, sectionHeadings) {
  let currentSection = null;
  
  for (let i = 0; i < sectionHeadings.length; i++) {
    const heading = sectionHeadings[i];
    
    // If cursor is after this heading start
    if (cursorPosition >= heading.start) {
      // Check if there's a next heading
      const nextHeading = sectionHeadings[i + 1];
      
      // If no next heading, or cursor is before next heading
      if (!nextHeading || cursorPosition < nextHeading.start) {
        currentSection = heading;
        break;
      }
    }
  }
  
  return currentSection;
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
