/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Set up event handlers
    document.getElementById("run").onclick = run;
    document.getElementById("getToc").onclick = getTableOfContents;
    document.getElementById("getSharePointPath").onclick = getSharePointPath;
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
      // Get all headings in the document
      const headings = context.document.body.paragraphs;
      context.load(headings, "text, styleBuiltIn, outlineLevel");
      
      await context.sync();

      const tocItems = [];
      
      for (let i = 0; i < headings.items.length; i++) {
        const paragraph = headings.items[i];
        
        // Check if this paragraph is a heading
        if (paragraph.styleBuiltIn && 
            (paragraph.styleBuiltIn.includes("Heading") || 
             paragraph.styleBuiltIn === "Title" ||
             paragraph.outlineLevel < 9)) {
          
          const level = getHeadingLevel(paragraph.styleBuiltIn, paragraph.outlineLevel);
          const text = paragraph.text.trim();
          
          if (text) {
            tocItems.push({
              text: text,
              level: level,
              style: paragraph.styleBuiltIn
            });
          }
        }
      }

      displayTableOfContents(tocItems);
      
    } catch (error) {
      console.error("Error getting table of contents:", error);
      document.getElementById("toc-content").innerHTML = 
        `<p style="color: red;">Error: ${error.message}</p>`;
      document.getElementById("toc-section").style.display = "block";
    }
  });
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
  if (styleBuiltIn) {
    if (styleBuiltIn === "Title") return 0;
    if (styleBuiltIn.includes("Heading")) {
      const match = styleBuiltIn.match(/Heading(\d+)/);
      return match ? parseInt(match[1]) : 1;
    }
  }
  return outlineLevel < 9 ? outlineLevel : 1;
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
          <div class="toc-item ${levelClass}" style="
            padding: 5px; 
            border-left: 3px solid ${getLevelColor(item.level)}; 
            background: ${index % 2 === 0 ? '#f9f9f9' : '#ffffff'};
            border-radius: 3px;
          ">
            <strong>H${item.level}:</strong> ${escapeHtml(item.text)}
            <br><small style="color: #666;">${item.style || 'Unknown style'}</small>
          </div>
        </li>
      `;
    });
    
    html += "</ul>";
    tocContainer.innerHTML = html;
  }
  
  tocSection.style.display = "block";
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
