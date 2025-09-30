/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// Global variables for tracking
let currentTocItems = [];
let autoTrackInterval = null;

// Enhanced logging function for debugging
function debugLog(functionName, message, data = null) {
  const timestamp = new Date().toISOString();
  const logMessage = `[${timestamp}] ${functionName}: ${message}`;
  
  console.log(logMessage);
  if (data) {
    console.log("Data:", data);
  }
  
  // Also display in UI for easier debugging
  const debugElement = document.getElementById("debug-info");
  if (debugElement) {
    const logEntry = document.createElement("div");
    logEntry.style.fontSize = "12px";
    logEntry.style.color = "#666";
    logEntry.style.marginBottom = "2px";
    logEntry.textContent = logMessage;
    debugElement.appendChild(logEntry);
    
    // Keep only last 10 log entries
    while (debugElement.children.length > 10) {
      debugElement.removeChild(debugElement.firstChild);
    }
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Add event listeners
    document.getElementById("getToc").onclick = () => {
      debugLog("getToc", "Button clicked - starting TOC extraction");
      getTableOfContents();
    };
    document.getElementById("testMinimal").onclick = () => {
      debugLog("testMinimal", "Button clicked - testing minimal API");
      testMinimal();
    };
    document.getElementById("getSharePointPath").onclick = () => {
      debugLog("getSharePointPath", "Button clicked - getting SharePoint path");
      getSharePointPath();
    };
    document.getElementById("getCurrentSection").onclick = () => {
      debugLog("getCurrentSection", "Button clicked - getting current section");
      const selectedStyle = document.getElementById("sectionHeadingStyle").value;
      getCurrentSection(selectedStyle);
    };
    document.getElementById("testApi").onclick = () => {
      debugLog("testApi", "Button clicked - testing AWS API");
      testAwsApi();
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

// Get Table of Contents - using content controls approach
export async function getTableOfContents() {
  return Word.run(async (context) => {
    try {
      console.log("Starting getTableOfContents with content controls...");
      
      // Clear previous results
      currentTocItems = [];
      document.getElementById("toc-content").innerHTML = "Loading...";
      document.getElementById("toc-section").style.display = "block";
      
      const tocItems = [];
      
      // Try content controls first
      try {
        console.log("Checking for content controls...");
        const contentControls = context.document.contentControls;
        context.load(contentControls, "items");
        await context.sync();
        
        console.log(`Found ${contentControls.items.length} content controls`);
        
        if (contentControls.items.length > 0) {
          // Load content control properties
          for (let i = 0; i < contentControls.items.length; i++) {
            const control = contentControls.items[i];
            context.load(control, "text, title, tag, type");
          }
          await context.sync();
          
          // Process content controls for headings
          for (let i = 0; i < contentControls.items.length; i++) {
            const control = contentControls.items[i];
            const text = control.text ? control.text.trim() : "";
            const title = control.title || "";
            const tag = control.tag || "";
            
            // Check if this content control represents a heading
            if (
              text &&
              (title.toLowerCase().includes("heading") ||
                tag.toLowerCase().includes("heading") ||
                title.toLowerCase().includes("title"))
            ) {
              const level = extractLevelFromControl(title, tag);
              
              tocItems.push({
                text: text,
                level: level,
                style: title || tag || "Content Control",
                index: tocItems.length,
                type: "contentControl",
              });
              
              console.log(`Found heading content control: "${text}" (${title || tag})`);
            }
          }
        }
      } catch (contentControlError) {
        console.warn("Content controls approach failed:", contentControlError);
      }
      
      // Fallback to paragraph scanning if no headings found in content controls
      if (tocItems.length === 0) {
        console.log("No headings found in content controls, falling back to paragraph scanning...");
        
        // Get all paragraphs
        const paragraphs = context.document.body.paragraphs;
        context.load(paragraphs, "items");
        await context.sync();
        
        console.log(`Found ${paragraphs.items.length} paragraphs total`);
        
        // Parse ALL paragraphs to find headings
        for (let i = 0; i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i];
          context.load(para, "text, styleBuiltIn");
        }
        await context.sync();
        
        // Process all paragraphs and find headings
        for (let i = 0; i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i];
          const text = para.text ? para.text.trim() : "";
          const style = para.styleBuiltIn ? para.styleBuiltIn.toString() : "";
          
          // Check if this paragraph is a heading
          if (text && text.length < 200 && (style.includes("Heading") || style === "Title")) {
            const level = getHeadingLevel(style);
            
            tocItems.push({
              text: text,
              level: level,
              style: style,
              index: tocItems.length,
              type: "paragraph",
            });
            
            console.log(`Found heading: "${text}" (${style})`);
          }
        }
      }
      
      console.log(`Found ${tocItems.length} headings total`);
      
      // Store and display results
      currentTocItems = tocItems;
      displayTableOfContents(currentTocItems);
      
    } catch (error) {
      console.error("Error getting table of contents:", error);
      document.getElementById("toc-content").innerHTML = `<p style="color: red;">Error: ${error.message}</p>`;
      document.getElementById("toc-section").style.display = "block";
    }
  });
}

// Simple helper to get heading level from style
function getHeadingLevel(style) {
  if (style === "Title") return 0;
  if (style.includes("Heading")) {
    const match = style.match(/(\d+)/);
    return match ? parseInt(match[1]) : 1;
  }
  return 1;
}

// Helper to extract heading level from content control title or tag
function extractLevelFromControl(title, tag) {
  const text = (title + " " + tag).toLowerCase();
  
  if (text.includes("title")) return 0;
  
  // Look for heading numbers
  const match = text.match(/heading\s*(\d+)|h(\d+)|level\s*(\d+)/);
  if (match) {
    const level = parseInt(match[1] || match[2] || match[3]);
    return level && level > 0 && level < 10 ? level : 1;
  }
  
  return 1;
}

// Helper function to scan for heading styles efficiently - NO document parsing
async function scanForHeadingStyles(context, tocItems) {
  try {
    console.log("Scanning for heading styles using direct style queries...");

    // Define the heading styles we want to find
    const headingStyles = ["Title", "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5", "Heading 6"];

    // For each heading style, search directly for that style
    for (const styleName of headingStyles) {
      try {
        console.log(`Searching for style: ${styleName}`);

        // Use the style-based search instead of wildcard
        const styleResults = context.document.body.search("", {
          matchCase: false,
          matchWholeWord: false,
          matchWildcards: false,
          // This searches for paragraphs with the specific style
          matchStyle: styleName,
        });

        context.load(styleResults, "items");
        await context.sync();

        console.log(`Found ${styleResults.items.length} items with style ${styleName}`);

        // Process each result
        for (const result of styleResults.items) {
          context.load(result, "text");
        }
        await context.sync();

        for (const result of styleResults.items) {
          const text = result.text ? result.text.trim() : "";
          if (text && text.length < 300) {
            // Reasonable heading length
            const level = getHeadingLevelFromStyleName(styleName);

            tocItems.push({
              text: text,
              level: level,
              style: styleName,
              index: tocItems.length,
            });

            console.log(`Found ${styleName}: "${text}"`);
          }
        }
      } catch (styleError) {
        console.warn(`Could not search for style ${styleName}:`, styleError);
      }
    }
  } catch (error) {
    console.error("Error in scanForHeadingStyles:", error);
  }
}

// Helper to get heading level from style name
function getHeadingLevelFromStyleName(styleName) {
  if (styleName === "Title") return 0;
  if (styleName.includes("Heading")) {
    const match = styleName.match(/(\d+)/);
    return match ? parseInt(match[1]) : 1;
  }
  return 1;
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

      document.getElementById("toc-content").innerHTML = `<p style="color: green;">✅ Minimal test passed!</p>
         <p>Document has ${body.text.length} characters</p>`;
    } catch (error) {
      console.error("Minimal test failed:", error);
      document.getElementById(
        "toc-content"
      ).innerHTML = `<p style="color: red;">❌ Minimal test failed: ${error.message}</p>`;
    }
  });
}

// Test AWS API function - GetDraftModelCollapse
export async function testAwsApi() {
  try {
    console.log("Testing AWS API - /draft/getAll/ endpoint...");
    
    // Show the API section
    document.getElementById("api-section").style.display = "block";

    // Parameters for GetDraftModelCollapse
    const params = {
      key: "demo_cuvitru",
      base: "usercache",
    };

    // Updated endpoint: POST to /draft/getAll/
    const baseUrl = "https://k43riamgd3.execute-api.us-east-2.amazonaws.com";
    const endpoint = "/draft/getAll/";
    const apiUrl = `${baseUrl}${endpoint}`;

    console.log("Full API URL:", apiUrl);
    console.log("Request method: POST with JSON body");
    console.log("Request body:", JSON.stringify(params, null, 2));

    // Make the API call with POST method and JSON body
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json",
      },
      mode: "cors",
      body: JSON.stringify(params),
    });

    console.log("API Response status:", response.status);
    console.log("API Response headers:", response.headers);

    // Update status
    document.getElementById("api-status").textContent = `Status: ${response.status} ${response.statusText}`;

    // Try to get response content - only read once!
    let responseData;
    const contentType = response.headers.get("content-type");

    try {
      if (contentType && contentType.includes("application/json")) {
        // Content-Type indicates JSON
        responseData = await response.json();
        document.getElementById("api-response").textContent = JSON.stringify(responseData, null, 2);
      } else {
        // Not JSON or unknown content type - get as text
        responseData = await response.text();
        document.getElementById("api-response").textContent = responseData;
      }

      // If successful and we got JSON data, show a summary
      if (response.ok && typeof responseData === "object") {
        console.log("Draft model retrieved successfully:", responseData);

        // Add some summary information
        let summary = `✅ Draft model retrieved successfully!\n\n`;
        summary += `Parameters used:\n- Key: ${params.key}\n- Base: ${params.base}\n\n`;
        summary += `Response data:\n${JSON.stringify(responseData, null, 2)}`;

        document.getElementById("api-response").textContent = summary;
      }
    } catch (parseError) {
      console.warn("Error parsing response:", parseError);
      // If parsing fails, try to get raw text (but response might already be consumed)
      try {
        responseData = await response.text();
        document.getElementById("api-response").textContent = `Response (parse error):\n${responseData}`;
      } catch (textError) {
        document.getElementById("api-response").textContent = `Error reading response: ${parseError.message}`;
      }
    }

    console.log("API Response data:", responseData);

    if (response.ok) {
      document.getElementById("api-status").style.color = "green";
    } else {
      document.getElementById("api-status").style.color = "orange";
    }
  } catch (error) {
    console.error("Error calling GetDraftModelCollapse API:", error);

    document.getElementById("api-section").style.display = "block";
    document.getElementById("api-status").textContent = `Error: ${error.message}`;
    document.getElementById("api-status").style.color = "red";
    document.getElementById(
      "api-response"
    ).textContent = `Error details:\n${error.name}: ${error.message}\n\nThis might be due to:\n- CORS policy restrictions\n- Network connectivity issues\n- API endpoint not available\n- Authentication required\n- Invalid parameters\n\nTried to call: GetDraftModelCollapse\nWith parameters:\n- key: d074010377c21837e01f22424d97cc6e\n- base: usercache`;
  }
}

// Simplified heading level function
function getHeadingLevelSimple(style) {
  if (style === "Title") return 0;
  if (style.includes("Heading")) {
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
        if (url.includes("sharepoint") || url.includes(".sharepoint.com")) {
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

// Helper function to determine heading level (extended)
function getHeadingLevelExtended(styleBuiltIn, outlineLevel) {
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
            background: ${index % 2 === 0 ? "#f9f9f9" : "#ffffff"};
            border-radius: 3px;
            cursor: pointer;
            transition: all 0.2s ease;
          ">
            <strong>H${item.level}:</strong> ${escapeHtml(item.text)}
            <br><small style="color: #666;">${item.style || "Unknown style"}</small>
          </div>
        </li>
      `;
    });

    html += "</ul>";
    tocContainer.innerHTML = html;

    // Add click handlers to TOC items for navigation
    const tocItemElements = tocContainer.querySelectorAll(".toc-item");
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
  if (documentUrl && documentUrl.includes("sharepoint")) {
    pathElement.textContent = documentUrl;

    // Extract filename from URL
    const urlParts = documentUrl.split("/");
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

// Get Current Section - finds which section the cursor is in based on content controls or heading boundaries
export async function getCurrentSection(sectionHeadingStyle = "Heading 1") {
  return Word.run(async (context) => {
    try {
      console.log(`Finding current section using content controls and ${sectionHeadingStyle} boundaries...`);

      // Show position section
      document.getElementById("position-section").style.display = "block";
      document.getElementById("current-section").textContent = "Detecting...";

      // First, check if cursor is inside a content control
      const selection = context.document.getSelection();
      context.load(selection, "start,parentContentControl");
      await context.sync();

      const ccs = selection.parentContentControl;
      if (ccs && ccs.isNullObject === false) {
        context.load(ccs, "title,tag,id,text");
        await context.sync();
      }

      let currentSection = null;
      const cursorPosition = selection.start;
      console.log(`Cursor position: ${cursorPosition}`);

      // Check if we're inside a content control
      if (ccs && ccs.isNullObject === false) {
        console.log("Cursor is inside content control:", ccs.title, ccs.tag, ccs.id);
        
        // Check if this content control represents a section/heading
        const title = ccs.title || "";
        const tag = ccs.tag || "";
        const text = ccs.text || "";
        
        if (
          title.toLowerCase().includes("section") ||
          title.toLowerCase().includes("heading") ||
          tag.toLowerCase().includes("section") ||
          tag.toLowerCase().includes("heading") ||
          title.toLowerCase().includes(sectionHeadingStyle.toLowerCase())
        ) {
          currentSection = {
            title: text.trim() || title || tag || `Content Control ${ccs.id}`,
            start: cursorPosition,
            end: cursorPosition,
            style: `Content Control: ${title || tag}`,
            type: "contentControl",
            id: ccs.id,
          };
          
          console.log(`Found section content control: "${currentSection.title}"`);
        } else {
          console.log("Content control doesn't appear to be a section heading");
        }
      } else {
        console.log("Cursor is not inside any content control.");
      }

      // If no content control section found, fall back to paragraph-based detection
      if (!currentSection) {
        console.log(`Falling back to ${sectionHeadingStyle} paragraph detection...`);
        
        // Find all headings of the specified level (section boundaries)
        const sectionHeadings = await findHeadingsByStyle(context, sectionHeadingStyle);
        console.log(`Found ${sectionHeadings.length} section headings`);

        if (sectionHeadings.length === 0) {
          document.getElementById(
            "current-section"
          ).textContent = `No ${sectionHeadingStyle} or section content controls found`;
          document.getElementById("current-position").textContent = "Cannot determine section";
          document.getElementById(
            "cursor-info"
          ).textContent = `Add some ${sectionHeadingStyle} headings or section content controls to use this feature`;
          return;
        }

        // Find which section the cursor is in
        currentSection = findSectionFromPosition(cursorPosition, sectionHeadings);
      }

      // Update display
      if (currentSection) {
        const sectionType = currentSection.type === "contentControl" ? "Content Control" : "Heading";
        document.getElementById("current-section").textContent = `Section: ${currentSection.title} (${sectionType})`;
        document.getElementById(
          "current-position"
        ).textContent = `Position: ${cursorPosition} (in ${currentSection.title})`;
        
        if (currentSection.type === "contentControl") {
          document.getElementById("cursor-info").textContent = `Inside content control: ${currentSection.style}`;
        } else {
          document.getElementById("cursor-info").textContent = `Section starts at position ${currentSection.start}`;
        }
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
      const style = paragraph.styleBuiltIn ? paragraph.styleBuiltIn.toString() : "";

      if (
        style === headingStyle ||
        (headingStyle === "Heading 1" && style.includes("Heading1")) ||
        (headingStyle === "Heading 2" && style.includes("Heading2"))
      ) {
        const text = paragraph.text ? paragraph.text.trim() : "";
        if (text) {
          // Get the range to find position
          const range = paragraph.getRange();
          context.load(range, "start, end");
          await context.sync();

          headings.push({
            title: text,
            start: range.start,
            end: range.end,
            style: style,
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
          end: item.end,
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
    cursorInfo += ` (Selected: "${selectedText.substring(0, 50)}${selectedText.length > 50 ? "..." : ""}")`;
  }
  document.getElementById("cursor-info").textContent = cursorInfo;
}

// Highlight current section in TOC
function highlightCurrentSectionInTOC(currentSection) {
  // Remove previous highlights
  const tocItems = document.querySelectorAll(".toc-item");
  tocItems.forEach((item) => {
    item.classList.remove("current-section");
    item.style.backgroundColor = "";
    item.style.fontWeight = "";
  });

  // Highlight current section
  if (currentSection && currentSection.index >= 0) {
    const currentItem = tocItems[currentSection.index];
    if (currentItem) {
      currentItem.classList.add("current-section");
      currentItem.style.backgroundColor = "#fff4ce";
      currentItem.style.fontWeight = "bold";
      currentItem.style.border = "2px solid #ffb900";

      // Scroll to current item in TOC
      currentItem.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }
  }
}

// Helper functions
function getLevelColor(level) {
  const colors = ["#0078d4", "#107c10", "#d83b01", "#b146c2", "#00bcf2", "#8764b8"];
  return colors[level % colors.length];
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}
