// src/googleDocsApiHelpers.ts
import { google, docs_v1 } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import { UserError } from 'fastmcp';
import { TextStyleArgs, ParagraphStyleArgs, hexToRgbColor, NotImplementedError } from './types.js';

type Docs = docs_v1.Docs; // Alias for convenience

// --- Constants ---
const MAX_BATCH_UPDATE_REQUESTS = 50; // Google API limits batch size

// --- Core Helper to Execute Batch Updates ---
export async function executeBatchUpdate(docs: Docs, documentId: string, requests: docs_v1.Schema$Request[]): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
if (!requests || requests.length === 0) {
// console.warn("executeBatchUpdate called with no requests.");
return {}; // Nothing to do
}

    // TODO: Consider splitting large request arrays into multiple batches if needed
    if (requests.length > MAX_BATCH_UPDATE_REQUESTS) {
         // console.warn(`Attempting batch update with ${requests.length} requests, exceeding typical limits. May fail.`);
    }

    try {
        const response = await docs.documents.batchUpdate({
            documentId: documentId,
            requestBody: { requests },
        });
        return response.data;
    } catch (error: any) {
        // console.error(`Google API batchUpdate Error for doc ${documentId}:`, error.response?.data || error.message);
        // Translate common API errors to UserErrors
        if (error.code === 400 && error.message.includes('Invalid requests')) {
             // Try to extract more specific info if available
             const details = error.response?.data?.error?.details;
             let detailMsg = '';
             if (details && Array.isArray(details)) {
                 detailMsg = details.map(d => d.description || JSON.stringify(d)).join('; ');
             }
            throw new UserError(`Invalid request sent to Google Docs API. Details: ${detailMsg || error.message}`);
        }
        if (error.code === 404) throw new UserError(`Document not found (ID: ${documentId}). Check the ID.`);
        if (error.code === 403) throw new UserError(`Permission denied for document (ID: ${documentId}). Ensure the authenticated user has edit access.`);
        // Generic internal error for others
        throw new Error(`Google API Error (${error.code}): ${error.message}`);
    }

}

// --- Text Finding Helper ---
// Robust text reconstruction with precise phrase matching and apostrophe handling
export async function findTextRange(docs: Docs, documentId: string, textToFind: string, instance: number = 1): Promise<{ startIndex: number; endIndex: number } | null> {
    try {
        // Get comprehensive document structure
        const res = await docs.documents.get({
            documentId,
            fields: 'body(content(paragraph(elements(startIndex,endIndex,textRun(content))),table(tableRows(tableCells(content(paragraph(elements(startIndex,endIndex,textRun(content))))))),startIndex,endIndex))',
        });

        if (!res.data.body?.content) {
            return null;
        }

        // Text segment with precise position mapping
        interface TextSegment {
            text: string;
            startIndex: number;
            endIndex: number;
        }

        // Extract all text segments in document order
        const textSegments: TextSegment[] = [];
        
        const extractTextFromContent = (content: any[]) => {
            for (const element of content) {
                if (element.paragraph?.elements) {
                    for (const pe of element.paragraph.elements) {
                        if (pe.textRun?.content && pe.startIndex !== undefined && pe.endIndex !== undefined) {
                            textSegments.push({
                                text: pe.textRun.content,
                                startIndex: pe.startIndex,
                                endIndex: pe.endIndex
                            });
                        }
                    }
                }
                
                if (element.table?.tableRows) {
                    for (const row of element.table.tableRows) {
                        if (row.tableCells) {
                            for (const cell of row.tableCells) {
                                if (cell.content) {
                                    extractTextFromContent(cell.content);
                                }
                            }
                        }
                    }
                }
            }
        };

        extractTextFromContent(res.data.body.content);
        
        // Sort by document position
        textSegments.sort((a, b) => a.startIndex - b.startIndex);

        // Build complete text reconstruction
        let fullText = '';
        const charToSegmentMap: Array<{ segmentIndex: number; charIndex: number; originalIndex: number }> = [];
        
        for (let segIndex = 0; segIndex < textSegments.length; segIndex++) {
            const segment = textSegments[segIndex];
            const segmentText = segment.text;
            
            for (let charIndex = 0; charIndex < segmentText.length; charIndex++) {
                charToSegmentMap.push({
                    segmentIndex: segIndex,
                    charIndex: charIndex,
                    originalIndex: segment.startIndex + charIndex
                });
                fullText += segmentText[charIndex];
            }
        }

        // Normalize text for apostrophe handling
        const normalizeText = (text: string): string => {
            return text
                .replace(/[\u0027\u2019\u2018\u201B\u0060\u00B4]/g, "'")  // Normalize apostrophes
                .replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"')  // Normalize quotes
                .replace(/[\u2013\u2014\u2015]/g, '-')  // Normalize dashes
                .replace(/\s+/g, ' ')   // Normalize whitespace
                .replace(/\u00A0/g, ' ') // Non-breaking space
                .replace(/\u2000-\u200A/g, ' ') // Unicode spaces
                .trim();
        };

        // Multi-level search strategy with strict validation
        const searchVariants = [
            textToFind,  // Exact search
            normalizeText(textToFind),  // Normalized apostrophes/quotes
            textToFind.replace(/[\u0027\u2019\u2018\u201B\u0060\u00B4]/g, "'"), // Standardize apostrophes
        ];

        // Helper function to find matches in reconstructed text and validate them
        const findValidatedMatches = (searchText: string): Array<{ startIndex: number; endIndex: number }> => {
            const matches: Array<{ startIndex: number; endIndex: number }> = [];
            
            // Try exact search first
            let searchIndex = 0;
            while (true) {
                const foundIndex = fullText.indexOf(searchText, searchIndex);
                if (foundIndex === -1) break;
                
                // Map to original indices using the char mapping
                const startMapping = charToSegmentMap[foundIndex];
                const endMapping = charToSegmentMap[foundIndex + searchText.length - 1];
                
                if (startMapping && endMapping) {
                    matches.push({
                        startIndex: startMapping.originalIndex,
                        endIndex: endMapping.originalIndex + 1
                    });
                }
                
                searchIndex = foundIndex + 1;
            }
            
            // If no exact matches, try normalized search
            if (matches.length === 0) {
                const normalizedFullText = normalizeText(fullText);
                const normalizedSearchText = normalizeText(searchText);
                
                searchIndex = 0;
                while (true) {
                    const foundIndex = normalizedFullText.indexOf(normalizedSearchText, searchIndex);
                    if (foundIndex === -1) break;
                    
                    // Map to original indices - this is trickier with normalization
                    // We need to find the corresponding position in the original text
                    let originalFoundIndex = 0;
                    let normalizedIndex = 0;
                    
                    // Count characters until we reach the found position in normalized text
                    while (normalizedIndex < foundIndex && originalFoundIndex < fullText.length) {
                        const originalChar = fullText[originalFoundIndex];
                        const normalizedChar = normalizeText(originalChar);
                        
                        if (normalizedChar.length > 0) {
                            normalizedIndex += normalizedChar.length;
                        }
                        originalFoundIndex++;
                    }
                    
                    // Calculate end position
                    let originalEndIndex = originalFoundIndex;
                    let remainingNormalizedLength = normalizedSearchText.length;
                    
                    while (remainingNormalizedLength > 0 && originalEndIndex < fullText.length) {
                        const originalChar = fullText[originalEndIndex];
                        const normalizedChar = normalizeText(originalChar);
                        
                        if (normalizedChar.length > 0) {
                            remainingNormalizedLength -= normalizedChar.length;
                        }
                        originalEndIndex++;
                    }
                    
                    // Map to document indices
                    const startMapping = charToSegmentMap[originalFoundIndex];
                    const endMapping = charToSegmentMap[originalEndIndex - 1];
                    
                    if (startMapping && endMapping) {
                        // Validate by extracting the actual text
                        const actualText = extractTextFromRange(startMapping.originalIndex, endMapping.originalIndex + 1);
                        if (normalizeText(actualText) === normalizedSearchText) {
                            matches.push({
                                startIndex: startMapping.originalIndex,
                                endIndex: endMapping.originalIndex + 1
                            });
                        }
                    }
                    
                    searchIndex = foundIndex + 1;
                }
            }
            
            return matches;
        };

        // Helper function to extract actual text from document range for validation
        const extractTextFromRange = (startIndex: number, endIndex: number): string => {
            let result = '';
            for (const segment of textSegments) {
                const segStart = segment.startIndex;
                const segEnd = segment.endIndex;
                
                // Check if this segment overlaps with our range
                if (segStart < endIndex && segEnd > startIndex) {
                    const overlapStart = Math.max(segStart, startIndex);
                    const overlapEnd = Math.min(segEnd, endIndex);
                    const textStart = overlapStart - segStart;
                    const textEnd = overlapEnd - segStart;
                    result += segment.text.substring(textStart, textEnd);
                }
            }
            return result;
        };

        // Try each search variant
        let allMatches: Array<{ startIndex: number; endIndex: number }> = [];
        
        for (const searchVariant of searchVariants) {
            const matches = findValidatedMatches(searchVariant);
            allMatches = allMatches.concat(matches);
            
            // If we found enough matches, stop searching
            if (allMatches.length >= instance) {
                break;
            }
        }

        // Remove duplicates and sort by position
        const uniqueMatches = Array.from(
            new Map(allMatches.map(m => [`${m.startIndex}-${m.endIndex}`, m])).values()
        );
        uniqueMatches.sort((a, b) => a.startIndex - b.startIndex);

        // Return the requested instance
        if (uniqueMatches.length >= instance) {
            return uniqueMatches[instance - 1];
        }

        return null;

    } catch (error: any) {
        if (error.code === 404) throw new UserError(`Document not found while searching text (ID: ${documentId}).`);
        if (error.code === 403) throw new UserError(`Permission denied while searching text in doc ${documentId}.`);
        throw new Error(`Failed to retrieve doc for text searching: ${error.message || 'Unknown error'}`);
    }
}

// Split-and-highlight approach for apostrophe-containing phrases
export async function highlightApostrophePhrase(
    docs: Docs, 
    documentId: string, 
    textToFind: string, 
    instance: number, 
    styleParams: TextStyleArgs
): Promise<boolean> {
    try {
        // Common apostrophe patterns that cause issues
        const apostrophePatterns = [
            { regex: /(\w+)[\u0027\u2019\u2018\u201B\u0060\u00B4](s?\s+.+)/i, description: "word's phrase..." },
            { regex: /(\w+)[\u0027\u2019\u2018\u201B\u0060\u00B4](re?\s+.+)/i, description: "word're phrase..." },
            { regex: /(\w+)[\u0027\u2019\u2018\u201B\u0060\u00B4](ll?\s+.+)/i, description: "word'll phrase..." },
            { regex: /(\w+)[\u0027\u2019\u2018\u201B\u0060\u00B4](ve?\s+.+)/i, description: "word've phrase..." },
            { regex: /(\w+)[\u0027\u2019\u2018\u201B\u0060\u00B4](d?\s+.+)/i, description: "word'd phrase..." },
            { regex: /(\w+)[\u0027\u2019\u2018\u201B\u0060\u00B4](t?\s+.+)/i, description: "word't phrase..." }
        ];

        for (const pattern of apostrophePatterns) {
            const match = textToFind.match(pattern.regex);
            if (match) {
                const beforeApostrophe = match[1]; // e.g., "that"
                const afterApostrophe = match[2];  // e.g., "s almost 20% more"
                
                let successfulHighlights = 0;
                
                // Try to find and highlight the beforeApostrophe part
                const beforeRange = await findTextRange(docs, documentId, beforeApostrophe, instance);
                if (beforeRange) {
                    const beforeRequest = buildUpdateTextStyleRequest(beforeRange.startIndex, beforeRange.endIndex, styleParams);
                    if (beforeRequest) {
                        await executeBatchUpdate(docs, documentId, [beforeRequest.request]);
                        successfulHighlights++;
                    }
                }
                
                // Try to find and highlight the afterApostrophe part
                const afterRange = await findTextRange(docs, documentId, afterApostrophe, instance);
                if (afterRange) {
                    const afterRequest = buildUpdateTextStyleRequest(afterRange.startIndex, afterRange.endIndex, styleParams);
                    if (afterRequest) {
                        await executeBatchUpdate(docs, documentId, [afterRequest.request]);
                        successfulHighlights++;
                    }
                }
                
                // Consider it successful if we highlighted both parts
                return successfulHighlights >= 2;
            }
        }
        
        // If no apostrophe pattern matched, fall back to standard search
        const range = await findTextRange(docs, documentId, textToFind, instance);
        if (range) {
            const request = buildUpdateTextStyleRequest(range.startIndex, range.endIndex, styleParams);
            if (request) {
                await executeBatchUpdate(docs, documentId, [request.request]);
                return true;
            }
        }
        
        return false;
        
    } catch (error: any) {
        console.error(`Error in highlightApostrophePhrase: ${error.message}`);
        return false;
    }
}


// --- Paragraph Boundary Helper ---
// Enhanced version to handle document structural elements more robustly
export async function getParagraphRange(docs: Docs, documentId: string, indexWithin: number): Promise<{ startIndex: number; endIndex: number } | null> {
try {
    // console.log(`Finding paragraph containing index ${indexWithin} in document ${documentId}`);
    
    // Request more detailed document structure to handle nested elements
    const res = await docs.documents.get({
        documentId,
        // Request more comprehensive structure information
        fields: 'body(content(startIndex,endIndex,paragraph,table,sectionBreak,tableOfContents))',
    });

    if (!res.data.body?.content) {
        // console.warn(`No content found in document ${documentId}`);
        return null;
    }

    // Find paragraph containing the index
    // We'll look at all structural elements recursively
    const findParagraphInContent = (content: any[]): { startIndex: number; endIndex: number } | null => {
        for (const element of content) {
            // Check if we have element boundaries defined
            if (element.startIndex !== undefined && element.endIndex !== undefined) {
                // Check if index is within this element's range first
                if (indexWithin >= element.startIndex && indexWithin < element.endIndex) {
                    // If it's a paragraph, we've found our target
                    if (element.paragraph) {
                        // console.log(`Found paragraph containing index ${indexWithin}, range: ${element.startIndex}-${element.endIndex}`);
                        return { 
                            startIndex: element.startIndex, 
                            endIndex: element.endIndex 
                        };
                    }
                    
                    // If it's a table, we need to check cells recursively
                    if (element.table && element.table.tableRows) {
                        // console.log(`Index ${indexWithin} is within a table, searching cells...`);
                        for (const row of element.table.tableRows) {
                            if (row.tableCells) {
                                for (const cell of row.tableCells) {
                                    if (cell.content) {
                                        const result = findParagraphInContent(cell.content);
                                        if (result) return result;
                                    }
                                }
                            }
                        }
                    }
                    
                    // For other structural elements, we didn't find a paragraph
                    // but we know the index is within this element
                    // console.warn(`Index ${indexWithin} is within element (${element.startIndex}-${element.endIndex}) but not in a paragraph`);
                }
            }
        }
        
        return null;
    };

    const paragraphRange = findParagraphInContent(res.data.body.content);
    
    if (!paragraphRange) {
        // console.warn(`Could not find paragraph containing index ${indexWithin}`);
    } else {
        // console.log(`Returning paragraph range: ${paragraphRange.startIndex}-${paragraphRange.endIndex}`);
    }
    
    return paragraphRange;

} catch (error: any) {
    // console.error(`Error getting paragraph range for index ${indexWithin} in doc ${documentId}: ${error.message || 'Unknown error'}`);
    if (error.code === 404) throw new UserError(`Document not found while finding paragraph (ID: ${documentId}).`);
    if (error.code === 403) throw new UserError(`Permission denied while accessing doc ${documentId}.`);
    throw new Error(`Failed to find paragraph: ${error.message || 'Unknown error'}`);
}
}

// --- Style Request Builders ---

export function buildUpdateTextStyleRequest(
startIndex: number,
endIndex: number,
style: TextStyleArgs
): { request: docs_v1.Schema$Request, fields: string[] } | null {
    const textStyle: docs_v1.Schema$TextStyle = {};
const fieldsToUpdate: string[] = [];

    if (style.bold !== undefined) { textStyle.bold = style.bold; fieldsToUpdate.push('bold'); }
    if (style.italic !== undefined) { textStyle.italic = style.italic; fieldsToUpdate.push('italic'); }
    if (style.underline !== undefined) { textStyle.underline = style.underline; fieldsToUpdate.push('underline'); }
    if (style.strikethrough !== undefined) { textStyle.strikethrough = style.strikethrough; fieldsToUpdate.push('strikethrough'); }
    if (style.fontSize !== undefined) { textStyle.fontSize = { magnitude: style.fontSize, unit: 'PT' }; fieldsToUpdate.push('fontSize'); }
    if (style.fontFamily !== undefined) { textStyle.weightedFontFamily = { fontFamily: style.fontFamily }; fieldsToUpdate.push('weightedFontFamily'); }
    if (style.foregroundColor !== undefined) {
        const rgbColor = hexToRgbColor(style.foregroundColor);
        if (!rgbColor) throw new UserError(`Invalid foreground hex color format: ${style.foregroundColor}`);
        textStyle.foregroundColor = { color: { rgbColor: rgbColor } }; fieldsToUpdate.push('foregroundColor');
    }
     if (style.backgroundColor !== undefined) {
        const rgbColor = hexToRgbColor(style.backgroundColor);
        if (!rgbColor) throw new UserError(`Invalid background hex color format: ${style.backgroundColor}`);
        textStyle.backgroundColor = { color: { rgbColor: rgbColor } }; fieldsToUpdate.push('backgroundColor');
    }
    if (style.linkUrl !== undefined) {
        textStyle.link = { url: style.linkUrl }; fieldsToUpdate.push('link');
    }
    // TODO: Handle clearing formatting

    if (fieldsToUpdate.length === 0) return null; // No styles to apply

    const request: docs_v1.Schema$Request = {
        updateTextStyle: {
            range: { startIndex, endIndex },
            textStyle: textStyle,
            fields: fieldsToUpdate.join(','),
        }
    };
    return { request, fields: fieldsToUpdate };

}

export function buildUpdateParagraphStyleRequest(
startIndex: number,
endIndex: number,
style: ParagraphStyleArgs
): { request: docs_v1.Schema$Request, fields: string[] } | null {
    // Create style object and track which fields to update
    const paragraphStyle: docs_v1.Schema$ParagraphStyle = {};
    const fieldsToUpdate: string[] = [];

    // console.log(`Building paragraph style request for range ${startIndex}-${endIndex} with options:`, style);

    // Process alignment option (LEFT, CENTER, RIGHT, JUSTIFIED)
    if (style.alignment !== undefined) { 
        paragraphStyle.alignment = style.alignment; 
        fieldsToUpdate.push('alignment'); 
        // console.log(`Setting alignment to ${style.alignment}`);
    }
    
    // Process indentation options
    if (style.indentStart !== undefined) { 
        paragraphStyle.indentStart = { magnitude: style.indentStart, unit: 'PT' }; 
        fieldsToUpdate.push('indentStart'); 
        // console.log(`Setting left indent to ${style.indentStart}pt`);
    }
    
    if (style.indentEnd !== undefined) { 
        paragraphStyle.indentEnd = { magnitude: style.indentEnd, unit: 'PT' }; 
        fieldsToUpdate.push('indentEnd'); 
        // console.log(`Setting right indent to ${style.indentEnd}pt`);
    }
    
    // Process spacing options
    if (style.spaceAbove !== undefined) { 
        paragraphStyle.spaceAbove = { magnitude: style.spaceAbove, unit: 'PT' }; 
        fieldsToUpdate.push('spaceAbove'); 
        // console.log(`Setting space above to ${style.spaceAbove}pt`);
    }
    
    if (style.spaceBelow !== undefined) { 
        paragraphStyle.spaceBelow = { magnitude: style.spaceBelow, unit: 'PT' }; 
        fieldsToUpdate.push('spaceBelow'); 
        // console.log(`Setting space below to ${style.spaceBelow}pt`);
    }
    
    // Process named style types (headings, etc.)
    if (style.namedStyleType !== undefined) { 
        paragraphStyle.namedStyleType = style.namedStyleType; 
        fieldsToUpdate.push('namedStyleType'); 
        // console.log(`Setting named style to ${style.namedStyleType}`);
    }
    
    // Process page break control
    if (style.keepWithNext !== undefined) { 
        paragraphStyle.keepWithNext = style.keepWithNext; 
        fieldsToUpdate.push('keepWithNext'); 
        // console.log(`Setting keepWithNext to ${style.keepWithNext}`);
    }

    // Verify we have styles to apply
    if (fieldsToUpdate.length === 0) {
        // console.warn("No paragraph styling options were provided");
        return null; // No styles to apply
    }

    // Build the request object
    const request: docs_v1.Schema$Request = {
        updateParagraphStyle: {
            range: { startIndex, endIndex },
            paragraphStyle: paragraphStyle,
            fields: fieldsToUpdate.join(','),
        }
    };
    
    // console.log(`Created paragraph style request with fields: ${fieldsToUpdate.join(', ')}`);
    return { request, fields: fieldsToUpdate };
}

// --- Specific Feature Helpers ---

export async function createTable(docs: Docs, documentId: string, rows: number, columns: number, index: number): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
    if (rows < 1 || columns < 1) {
        throw new UserError("Table must have at least 1 row and 1 column.");
    }
    const request: docs_v1.Schema$Request = {
insertTable: {
location: { index },
rows: rows,
columns: columns,
}
};
return executeBatchUpdate(docs, documentId, [request]);
}

export async function insertText(docs: Docs, documentId: string, text: string, index: number): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
    if (!text) return {}; // Nothing to insert
    const request: docs_v1.Schema$Request = {
insertText: {
location: { index },
text: text,
}
};
return executeBatchUpdate(docs, documentId, [request]);
}

// --- Complex / Stubbed Helpers ---

export async function findParagraphsMatchingStyle(
docs: Docs,
documentId: string,
styleCriteria: any // Define a proper type for criteria (e.g., { fontFamily: 'Arial', bold: true })
): Promise<{ startIndex: number; endIndex: number }[]> {
// TODO: Implement logic
// 1. Get document content with paragraph elements and their styles.
// 2. Iterate through paragraphs.
// 3. For each paragraph, check if its computed style matches the criteria.
// 4. Return ranges of matching paragraphs.
// console.warn("findParagraphsMatchingStyle is not implemented.");
throw new NotImplementedError("Finding paragraphs by style criteria is not yet implemented.");
// return [];
}

export async function detectAndFormatLists(
docs: Docs,
documentId: string,
startIndex?: number,
endIndex?: number
): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
// TODO: Implement complex logic
// 1. Get document content (paragraphs, text runs) in the specified range (or whole doc).
// 2. Iterate through paragraphs.
// 3. Identify sequences of paragraphs starting with list-like markers (e.g., "-", "*", "1.", "a)").
// 4. Determine nesting levels based on indentation or marker patterns.
// 5. Generate CreateParagraphBulletsRequests for the identified sequences.
// 6. Potentially delete the original marker text.
// 7. Execute the batch update.
// console.warn("detectAndFormatLists is not implemented.");
throw new NotImplementedError("Automatic list detection and formatting is not yet implemented.");
// return {};
}

export async function addCommentHelper(authClient: OAuth2Client, documentId: string, text: string, startIndex: number, endIndex: number): Promise<void> {
// Initialize Drive API client with the authenticated client
const drive = google.drive({version: 'v3', auth: authClient});

try {
    // Create comment with anchor to specific text range
    const response = await drive.comments.create({
        fileId: documentId,
        requestBody: {
            content: text,
            anchor: JSON.stringify({
                'r': 'kix.TextAnchor', // Google Docs text anchor type
                'textAnchor': {
                    'textRange': {
                        'start': startIndex,
                        'end': endIndex
                    }
                }
            })
        },
        fields: 'id,content,anchor'
    });
    
    console.log(`Successfully added comment with ID: ${response.data.id}`);
} catch (error: any) {
    console.error(`Error adding comment: ${error.message}`);
    if (error.code === 404) throw new UserError(`Document not found for commenting (ID: ${documentId}).`);
    if (error.code === 403) throw new UserError(`Permission denied for adding comment to doc (ID: ${documentId}). Ensure the authenticated user has comment access.`);
    throw new Error(`Failed to add comment: ${error.message || 'Unknown error'}`);
}
}