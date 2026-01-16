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
         console.warn(`Attempting batch update with ${requests.length} requests, exceeding typical limits. May fail.`);
    }

    try {
        const response = await docs.documents.batchUpdate({
            documentId: documentId,
            requestBody: { requests },
        });
        return response.data;
    } catch (error: any) {
        console.error(`Google API batchUpdate Error for doc ${documentId}:`, error.response?.data || error.message);
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
// This improved version is more robust in handling various text structure scenarios
export async function findTextRange(docs: Docs, documentId: string, textToFind: string, instance: number = 1): Promise<{ startIndex: number; endIndex: number } | null> {
try {
    // Request more detailed information about the document structure
    const res = await docs.documents.get({
        documentId,
        // Request more fields to handle various container types (not just paragraphs)
        fields: 'body(content(paragraph(elements(startIndex,endIndex,textRun(content))),table,sectionBreak,tableOfContents,startIndex,endIndex))',
    });

    if (!res.data.body?.content) {
        console.warn(`No content found in document ${documentId}`);
        return null;
    }

    // More robust text collection and index tracking
    let fullText = '';
    const segments: { text: string, start: number, end: number }[] = [];

    // Process all content elements, including structural ones
    const collectTextFromContent = (content: any[]) => {
        content.forEach(element => {
            // Handle paragraph elements
            if (element.paragraph?.elements) {
                element.paragraph.elements.forEach((pe: any) => {
                    if (pe.textRun?.content && pe.startIndex !== undefined && pe.endIndex !== undefined) {
                        const content = pe.textRun.content;
                        fullText += content;
                        segments.push({
                            text: content,
                            start: pe.startIndex,
                            end: pe.endIndex
                        });
                    }
                });
            }

            // Handle table elements - this is simplified and might need expansion
            if (element.table && element.table.tableRows) {
                element.table.tableRows.forEach((row: any) => {
                    if (row.tableCells) {
                        row.tableCells.forEach((cell: any) => {
                            if (cell.content) {
                                collectTextFromContent(cell.content);
                            }
                        });
                    }
                });
            }

            // Add handling for other structural elements as needed
        });
    };

    collectTextFromContent(res.data.body.content);

    // Sort segments by starting position to ensure correct ordering
    segments.sort((a, b) => a.start - b.start);

    console.log(`Document ${documentId} contains ${segments.length} text segments and ${fullText.length} characters in total.`);

    // Find the specified instance of the text
    let startIndex = -1;
    let endIndex = -1;
    let foundCount = 0;
    let searchStartIndex = 0;

    while (foundCount < instance) {
        const currentIndex = fullText.indexOf(textToFind, searchStartIndex);
        if (currentIndex === -1) {
            console.log(`Search text "${textToFind}" not found for instance ${foundCount + 1} (requested: ${instance})`);
            break;
        }

        foundCount++;
        console.log(`Found instance ${foundCount} of "${textToFind}" at position ${currentIndex} in full text`);

        if (foundCount === instance) {
            const targetStartInFullText = currentIndex;
            const targetEndInFullText = currentIndex + textToFind.length;
            let currentPosInFullText = 0;

            console.log(`Target text range in full text: ${targetStartInFullText}-${targetEndInFullText}`);

            for (const seg of segments) {
                const segStartInFullText = currentPosInFullText;
                const segTextLength = seg.text.length;
                const segEndInFullText = segStartInFullText + segTextLength;

                // Map from reconstructed text position to actual document indices
                if (startIndex === -1 && targetStartInFullText >= segStartInFullText && targetStartInFullText < segEndInFullText) {
                    startIndex = seg.start + (targetStartInFullText - segStartInFullText);
                    console.log(`Mapped start to segment ${seg.start}-${seg.end}, position ${startIndex}`);
                }

                if (targetEndInFullText > segStartInFullText && targetEndInFullText <= segEndInFullText) {
                    endIndex = seg.start + (targetEndInFullText - segStartInFullText);
                    console.log(`Mapped end to segment ${seg.start}-${seg.end}, position ${endIndex}`);
                    break;
                }

                currentPosInFullText = segEndInFullText;
            }

            if (startIndex === -1 || endIndex === -1) {
                console.warn(`Failed to map text "${textToFind}" instance ${instance} to actual document indices`);
                // Reset and try next occurrence
                startIndex = -1;
                endIndex = -1;
                searchStartIndex = currentIndex + 1;
                foundCount--;
                continue;
            }

            console.log(`Successfully mapped "${textToFind}" to document range ${startIndex}-${endIndex}`);
            return { startIndex, endIndex };
        }

        // Prepare for next search iteration
        searchStartIndex = currentIndex + 1;
    }

    console.warn(`Could not find instance ${instance} of text "${textToFind}" in document ${documentId}`);
    return null; // Instance not found or mapping failed for all attempts
} catch (error: any) {
    console.error(`Error finding text "${textToFind}" in doc ${documentId}: ${error.message || 'Unknown error'}`);
    if (error.code === 404) throw new UserError(`Document not found while searching text (ID: ${documentId}).`);
    if (error.code === 403) throw new UserError(`Permission denied while searching text in doc ${documentId}.`);
    throw new Error(`Failed to retrieve doc for text searching: ${error.message || 'Unknown error'}`);
}
}

// --- Paragraph Boundary Helper ---
// Enhanced version to handle document structural elements more robustly
export async function getParagraphRange(docs: Docs, documentId: string, indexWithin: number): Promise<{ startIndex: number; endIndex: number } | null> {
try {
    console.log(`Finding paragraph containing index ${indexWithin} in document ${documentId}`);

    // Request more detailed document structure to handle nested elements
    const res = await docs.documents.get({
        documentId,
        // Request more comprehensive structure information
        fields: 'body(content(startIndex,endIndex,paragraph,table,sectionBreak,tableOfContents))',
    });

    if (!res.data.body?.content) {
        console.warn(`No content found in document ${documentId}`);
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
                        console.log(`Found paragraph containing index ${indexWithin}, range: ${element.startIndex}-${element.endIndex}`);
                        return {
                            startIndex: element.startIndex,
                            endIndex: element.endIndex
                        };
                    }

                    // If it's a table, we need to check cells recursively
                    if (element.table && element.table.tableRows) {
                        console.log(`Index ${indexWithin} is within a table, searching cells...`);
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
                    console.warn(`Index ${indexWithin} is within element (${element.startIndex}-${element.endIndex}) but not in a paragraph`);
                }
            }
        }

        return null;
    };

    const paragraphRange = findParagraphInContent(res.data.body.content);

    if (!paragraphRange) {
        console.warn(`Could not find paragraph containing index ${indexWithin}`);
    } else {
        console.log(`Returning paragraph range: ${paragraphRange.startIndex}-${paragraphRange.endIndex}`);
    }

    return paragraphRange;

} catch (error: any) {
    console.error(`Error getting paragraph range for index ${indexWithin} in doc ${documentId}: ${error.message || 'Unknown error'}`);
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

    console.log(`Building paragraph style request for range ${startIndex}-${endIndex} with options:`, style);

    // Process alignment option (LEFT, CENTER, RIGHT, JUSTIFIED)
    if (style.alignment !== undefined) {
        paragraphStyle.alignment = style.alignment;
        fieldsToUpdate.push('alignment');
        console.log(`Setting alignment to ${style.alignment}`);
    }

    // Process indentation options
    if (style.indentStart !== undefined) {
        paragraphStyle.indentStart = { magnitude: style.indentStart, unit: 'PT' };
        fieldsToUpdate.push('indentStart');
        console.log(`Setting left indent to ${style.indentStart}pt`);
    }

    if (style.indentEnd !== undefined) {
        paragraphStyle.indentEnd = { magnitude: style.indentEnd, unit: 'PT' };
        fieldsToUpdate.push('indentEnd');
        console.log(`Setting right indent to ${style.indentEnd}pt`);
    }

    // Process spacing options
    if (style.spaceAbove !== undefined) {
        paragraphStyle.spaceAbove = { magnitude: style.spaceAbove, unit: 'PT' };
        fieldsToUpdate.push('spaceAbove');
        console.log(`Setting space above to ${style.spaceAbove}pt`);
    }

    if (style.spaceBelow !== undefined) {
        paragraphStyle.spaceBelow = { magnitude: style.spaceBelow, unit: 'PT' };
        fieldsToUpdate.push('spaceBelow');
        console.log(`Setting space below to ${style.spaceBelow}pt`);
    }

    // Process named style types (headings, etc.)
    if (style.namedStyleType !== undefined) {
        paragraphStyle.namedStyleType = style.namedStyleType;
        fieldsToUpdate.push('namedStyleType');
        console.log(`Setting named style to ${style.namedStyleType}`);
    }

    // Process page break control
    if (style.keepWithNext !== undefined) {
        paragraphStyle.keepWithNext = style.keepWithNext;
        fieldsToUpdate.push('keepWithNext');
        console.log(`Setting keepWithNext to ${style.keepWithNext}`);
    }

    // Verify we have styles to apply
    if (fieldsToUpdate.length === 0) {
        console.warn("No paragraph styling options were provided");
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

    console.log(`Created paragraph style request with fields: ${fieldsToUpdate.join(', ')}`);
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
console.warn("findParagraphsMatchingStyle is not implemented.");
throw new NotImplementedError("Finding paragraphs by style criteria is not yet implemented.");
// return [];
}

export async function detectAndFormatLists(
docs: Docs,
documentId: string,
startIndex?: number,
endIndex?: number
): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
    console.log(`detectAndFormatLists called for doc ${documentId}, range: ${startIndex}-${endIndex}`);

    // Get document content to analyze paragraphs
    const res = await docs.documents.get({
        documentId,
        fields: 'body(content(startIndex,endIndex,paragraph(elements(textRun(content)))))',
    });

    if (!res.data.body?.content) {
        throw new UserError("Document has no content to format.");
    }

    const requests: docs_v1.Schema$Request[] = [];
    const processedRanges: { start: number; end: number; bulletType: string }[] = [];

    // Analyze each paragraph in the document
    for (const element of res.data.body.content) {
        if (!element.paragraph || element.startIndex === undefined || element.startIndex === null ||
            element.endIndex === undefined || element.endIndex === null) {
            continue;
        }

        const paragraphStart: number = element.startIndex;
        const paragraphEnd: number = element.endIndex;

        // Skip if outside specified range
        if (startIndex !== undefined && paragraphEnd <= startIndex) continue;
        if (endIndex !== undefined && paragraphStart >= endIndex) continue;

        // Get paragraph text content
        let paragraphText = '';
        element.paragraph.elements?.forEach((pe: any) => {
            if (pe.textRun?.content) {
                paragraphText += pe.textRun.content;
            }
        });

        // Check for list-like patterns at the start of the paragraph
        const trimmedText = paragraphText.trimStart();
        let bulletPreset: string | null = null;
        let markerLength = 0;

        // Check for bullet markers: -, *, •
        if (/^[-*•]\s/.test(trimmedText)) {
            bulletPreset = 'BULLET_DISC_CIRCLE_SQUARE';
            markerLength = trimmedText.match(/^[-*•]\s*/)?.[0].length || 0;
        }
        // Check for numbered list markers: 1. 2. 3. or 1) 2) 3)
        else if (/^\d+[.)]\s/.test(trimmedText)) {
            bulletPreset = 'NUMBERED_DECIMAL_NESTED';
            markerLength = trimmedText.match(/^\d+[.)]\s*/)?.[0].length || 0;
        }
        // Check for lettered list markers: a. b. c. or a) b) c)
        else if (/^[a-zA-Z][.)]\s/.test(trimmedText)) {
            bulletPreset = 'NUMBERED_UPPERALPHA_ALPHA_ROMAN';
            markerLength = trimmedText.match(/^[a-zA-Z][.)]\s*/)?.[0].length || 0;
        }

        if (bulletPreset) {
            const leadingSpaces = paragraphText.length - paragraphText.trimStart().length;

            processedRanges.push({
                start: paragraphStart,
                end: paragraphEnd,
                bulletType: bulletPreset
            });

            // First, create the bullet for this paragraph
            requests.push({
                createParagraphBullets: {
                    range: {
                        startIndex: paragraphStart,
                        endIndex: paragraphEnd
                    },
                    bulletPreset: bulletPreset as any
                }
            });
        }
    }

    if (requests.length === 0) {
        console.log("No list-like paragraphs found to format.");
        return {};
    }

    console.log(`Found ${requests.length} paragraphs to convert to lists.`);

    // Execute the batch update to create bullets
    const result = await executeBatchUpdate(docs, documentId, requests);

    // Now we need to delete the marker text from each paragraph
    // We need to do this in reverse order to maintain correct indices
    const deleteRequests: docs_v1.Schema$Request[] = [];

    // Re-fetch document to get updated indices after bullet creation
    const updatedDoc = await docs.documents.get({
        documentId,
        fields: 'body(content(startIndex,endIndex,paragraph(elements(textRun(content)))))',
    });

    if (updatedDoc.data.body?.content) {
        for (const element of updatedDoc.data.body.content) {
            if (!element.paragraph || element.startIndex === undefined || element.startIndex === null) continue;

            let paragraphText = '';
            let textStartIndex: number = element.startIndex;

            element.paragraph.elements?.forEach((pe: any) => {
                if (pe.textRun?.content && pe.startIndex !== undefined && pe.startIndex !== null) {
                    if (paragraphText === '') {
                        textStartIndex = pe.startIndex as number;
                    }
                    paragraphText += pe.textRun.content;
                }
            });

            const trimmedText = paragraphText.trimStart();
            const leadingSpaces = paragraphText.length - trimmedText.length;

            let markerMatch = trimmedText.match(/^([-*•]\s*|\d+[.)]\s*|[a-zA-Z][.)]\s*)/);
            if (markerMatch) {
                const markerLength = markerMatch[0].length;
                const deleteStart: number = textStartIndex + leadingSpaces;
                const deleteEnd: number = deleteStart + markerLength;

                deleteRequests.push({
                    deleteContentRange: {
                        range: {
                            startIndex: deleteStart,
                            endIndex: deleteEnd
                        }
                    }
                });
            }
        }
    }

    // Execute delete requests in reverse order (to maintain indices)
    if (deleteRequests.length > 0) {
        deleteRequests.reverse();
        await executeBatchUpdate(docs, documentId, deleteRequests);
        console.log(`Deleted ${deleteRequests.length} marker texts.`);
    }

    return result;
}

/**
 * Converts paragraphs in a range to a bullet list (without needing markers)
 * @param docs - Google Docs API client
 * @param documentId - The document ID
 * @param startIndex - Start of range
 * @param endIndex - End of range
 * @param bulletPreset - Type of bullet list (default: BULLET_DISC_CIRCLE_SQUARE)
 */
export async function createBulletListInRange(
    docs: Docs,
    documentId: string,
    startIndex: number,
    endIndex: number,
    bulletPreset: string = 'BULLET_DISC_CIRCLE_SQUARE'
): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
    console.log(`Creating bullet list in doc ${documentId}, range: ${startIndex}-${endIndex}, preset: ${bulletPreset}`);

    const request: docs_v1.Schema$Request = {
        createParagraphBullets: {
            range: {
                startIndex: startIndex,
                endIndex: endIndex
            },
            bulletPreset: bulletPreset as any
        }
    };

    return executeBatchUpdate(docs, documentId, [request]);
}

/**
 * Removes bullet formatting from paragraphs in a range
 */
export async function removeBulletListInRange(
    docs: Docs,
    documentId: string,
    startIndex: number,
    endIndex: number
): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
    console.log(`Removing bullet list in doc ${documentId}, range: ${startIndex}-${endIndex}`);

    const request: docs_v1.Schema$Request = {
        deleteParagraphBullets: {
            range: {
                startIndex: startIndex,
                endIndex: endIndex
            }
        }
    };

    return executeBatchUpdate(docs, documentId, [request]);
}

export async function addCommentHelper(docs: Docs, documentId: string, text: string, startIndex: number, endIndex: number): Promise<void> {
// NOTE: Adding comments typically requires the Google Drive API v3 and different scopes!
// 'https://www.googleapis.com/auth/drive' or more specific comment scopes.
// This helper is a placeholder assuming Drive API client (`drive`) is available and authorized.
/*
const drive = google.drive({version: 'v3', auth: authClient}); // Assuming authClient is available
await drive.comments.create({
fileId: documentId,
requestBody: {
content: text,
anchor: JSON.stringify({ // Anchor format might need verification
'type': 'workbook#textAnchor', // Or appropriate type for Docs
'refs': [{
'docRevisionId': 'head', // Or specific revision
'range': {
'start': startIndex,
'end': endIndex,
}
}]
})
},
fields: 'id'
});
*/
console.warn("addCommentHelper requires Google Drive API and is not implemented.");
throw new NotImplementedError("Adding comments requires Drive API setup and is not yet implemented.");
}

// --- Image Insertion Helpers ---

/**
 * Inserts an inline image into a document from a publicly accessible URL
 * @param docs - Google Docs API client
 * @param documentId - The document ID
 * @param imageUrl - Publicly accessible URL to the image
 * @param index - Position in the document where image should be inserted (1-based)
 * @param width - Optional width in points
 * @param height - Optional height in points
 * @returns Promise with batch update response
 */
export async function insertInlineImage(
    docs: Docs,
    documentId: string,
    imageUrl: string,
    index: number,
    width?: number,
    height?: number
): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
    // Validate URL format
    try {
        new URL(imageUrl);
    } catch (e) {
        throw new UserError(`Invalid image URL format: ${imageUrl}`);
    }

    // Build the insertInlineImage request
    const request: docs_v1.Schema$Request = {
        insertInlineImage: {
            location: { index },
            uri: imageUrl,
            ...(width && height && {
                objectSize: {
                    height: { magnitude: height, unit: 'PT' },
                    width: { magnitude: width, unit: 'PT' }
                }
            })
        }
    };

    return executeBatchUpdate(docs, documentId, [request]);
}

/**
 * Uploads a local image file to Google Drive and returns its public URL
 * @param drive - Google Drive API client
 * @param localFilePath - Path to the local image file
 * @param parentFolderId - Optional parent folder ID (defaults to root)
 * @returns Promise with the public webContentLink URL
 */
export async function uploadImageToDrive(
    drive: any, // drive_v3.Drive type
    localFilePath: string,
    parentFolderId?: string
): Promise<string> {
    const fs = await import('fs');
    const path = await import('path');

    // Verify file exists
    if (!fs.existsSync(localFilePath)) {
        throw new UserError(`Image file not found: ${localFilePath}`);
    }

    // Get file name and mime type
    const fileName = path.basename(localFilePath);
    const mimeTypeMap: { [key: string]: string } = {
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.png': 'image/png',
        '.gif': 'image/gif',
        '.bmp': 'image/bmp',
        '.webp': 'image/webp',
        '.svg': 'image/svg+xml'
    };

    const ext = path.extname(localFilePath).toLowerCase();
    const mimeType = mimeTypeMap[ext] || 'application/octet-stream';

    // Upload file to Drive
    const fileMetadata: any = {
        name: fileName,
        mimeType: mimeType
    };

    if (parentFolderId) {
        fileMetadata.parents = [parentFolderId];
    }

    const media = {
        mimeType: mimeType,
        body: fs.createReadStream(localFilePath)
    };

    const uploadResponse = await drive.files.create({
        requestBody: fileMetadata,
        media: media,
        fields: 'id,webViewLink,webContentLink'
    });

    const fileId = uploadResponse.data.id;
    if (!fileId) {
        throw new Error('Failed to upload image to Drive - no file ID returned');
    }

    // Make the file publicly readable
    await drive.permissions.create({
        fileId: fileId,
        requestBody: {
            role: 'reader',
            type: 'anyone'
        }
    });

    // Get the webContentLink
    const fileInfo = await drive.files.get({
        fileId: fileId,
        fields: 'webContentLink'
    });

    const webContentLink = fileInfo.data.webContentLink;
    if (!webContentLink) {
        throw new Error('Failed to get public URL for uploaded image');
    }

    return webContentLink;
}

// --- Tab Management Helpers ---

/**
 * Interface for a tab with hierarchy level information
 */
export interface TabWithLevel extends docs_v1.Schema$Tab {
    level: number;
}

/**
 * Recursively collect all tabs from a document in a flat list with hierarchy info
 * @param doc - The Google Doc document object
 * @returns Array of tabs with nesting level information
 */
export function getAllTabs(doc: docs_v1.Schema$Document): TabWithLevel[] {
    const allTabs: TabWithLevel[] = [];
    if (!doc.tabs || doc.tabs.length === 0) {
        return allTabs;
    }

    for (const tab of doc.tabs) {
        addCurrentAndChildTabs(tab, allTabs, 0);
    }
    return allTabs;
}

/**
 * Recursive helper to add tabs with their nesting level
 * @param tab - The tab to add
 * @param allTabs - The accumulator array
 * @param level - Current nesting level (0 for top-level)
 */
function addCurrentAndChildTabs(tab: docs_v1.Schema$Tab, allTabs: TabWithLevel[], level: number): void {
    allTabs.push({ ...tab, level });
    if (tab.childTabs && tab.childTabs.length > 0) {
        for (const childTab of tab.childTabs) {
            addCurrentAndChildTabs(childTab, allTabs, level + 1);
        }
    }
}

/**
 * Get the text length from a DocumentTab
 * @param documentTab - The DocumentTab object
 * @returns Total character count
 */
export function getTabTextLength(documentTab: docs_v1.Schema$DocumentTab | undefined): number {
    let totalLength = 0;

    if (!documentTab?.body?.content) {
        return 0;
    }

    documentTab.body.content.forEach((element: any) => {
        // Handle paragraphs
        if (element.paragraph?.elements) {
            element.paragraph.elements.forEach((pe: any) => {
                if (pe.textRun?.content) {
                    totalLength += pe.textRun.content.length;
                }
            });
        }

        // Handle tables
        if (element.table?.tableRows) {
            element.table.tableRows.forEach((row: any) => {
                row.tableCells?.forEach((cell: any) => {
                    cell.content?.forEach((cellElement: any) => {
                        cellElement.paragraph?.elements?.forEach((pe: any) => {
                            if (pe.textRun?.content) {
                                totalLength += pe.textRun.content.length;
                            }
                        });
                    });
                });
            });
        }
    });

    return totalLength;
}

/**
 * Find a specific tab by ID in a document (searches recursively through child tabs)
 * @param doc - The Google Doc document object
 * @param tabId - The tab ID to search for
 * @returns The tab object if found, null otherwise
 */
export function findTabById(doc: docs_v1.Schema$Document, tabId: string): docs_v1.Schema$Tab | null {
    if (!doc.tabs || doc.tabs.length === 0) {
        return null;
    }

    // Helper function to search through tabs recursively
    const searchTabs = (tabs: docs_v1.Schema$Tab[]): docs_v1.Schema$Tab | null => {
        for (const tab of tabs) {
            if (tab.tabProperties?.tabId === tabId) {
                return tab;
            }
            // Recursively search child tabs
            if (tab.childTabs && tab.childTabs.length > 0) {
                const found = searchTabs(tab.childTabs);
                if (found) return found;
            }
        }
        return null;
    };

    return searchTabs(doc.tabs);
}

// --- Table Cell Editing Helpers ---

/**
 * Interface for table cell range information
 */
export interface TableCellRange {
    cellStartIndex: number;
    cellEndIndex: number;
    contentStartIndex: number;
    contentEndIndex: number;
    hasContent: boolean;
}

/**
 * Finds a table element in the document at or after the specified index
 * @param docs - Google Docs API client
 * @param documentId - The document ID
 * @param tableStartIndex - The approximate starting index of the table
 * @returns The table element and its metadata, or null if not found
 */
export async function findTableAtIndex(
    docs: Docs,
    documentId: string,
    tableStartIndex: number
): Promise<{ table: any; startIndex: number; endIndex: number } | null> {
    try {
        const res = await docs.documents.get({
            documentId,
            fields: 'body(content(startIndex,endIndex,table))',
        });

        if (!res.data.body?.content) {
            return null;
        }

        // Find the table at or near the specified index
        for (const element of res.data.body.content) {
            const startIdx = element.startIndex;
            const endIdx = element.endIndex;
            if (element.table && startIdx != null && endIdx != null) {
                // Allow some tolerance in finding the table (within 10 positions)
                if (startIdx >= tableStartIndex - 10 && startIdx <= tableStartIndex + 10) {
                    return {
                        table: element.table,
                        startIndex: startIdx,
                        endIndex: endIdx
                    };
                }
                // Also match if the provided index is exactly the table start
                if (startIdx === tableStartIndex) {
                    return {
                        table: element.table,
                        startIndex: startIdx,
                        endIndex: endIdx
                    };
                }
            }
        }

        // If not found with tolerance, try finding any table after the index
        for (const element of res.data.body.content) {
            const startIdx = element.startIndex;
            const endIdx = element.endIndex;
            if (element.table && startIdx != null && endIdx != null) {
                if (startIdx >= tableStartIndex) {
                    return {
                        table: element.table,
                        startIndex: startIdx,
                        endIndex: endIdx
                    };
                }
            }
        }

        return null;
    } catch (error: any) {
        console.error(`Error finding table at index ${tableStartIndex} in doc ${documentId}: ${error.message}`);
        throw new Error(`Failed to find table: ${error.message}`);
    }
}

/**
 * Gets the content range of a specific table cell
 * @param docs - Google Docs API client
 * @param documentId - The document ID
 * @param tableStartIndex - The starting index of the table
 * @param rowIndex - Row index (0-based)
 * @param columnIndex - Column index (0-based)
 * @returns The cell range information, or null if not found
 */
export async function getTableCellRange(
    docs: Docs,
    documentId: string,
    tableStartIndex: number,
    rowIndex: number,
    columnIndex: number
): Promise<TableCellRange | null> {
    try {
        // Get full document structure to find table cells
        const res = await docs.documents.get({
            documentId,
            fields: 'body(content(startIndex,endIndex,table(tableRows(tableCells(startIndex,endIndex,content(startIndex,endIndex,paragraph(elements(startIndex,endIndex,textRun(content)))))))))',
        });

        if (!res.data.body?.content) {
            console.warn(`No content found in document ${documentId}`);
            return null;
        }

        // Find the table
        let targetTable: any = null;
        for (const element of res.data.body.content) {
            const startIdx = element.startIndex;
            if (element.table && startIdx != null) {
                // Match table by start index (with some tolerance)
                if (Math.abs(startIdx - tableStartIndex) <= 10) {
                    targetTable = element.table;
                    break;
                }
            }
        }

        if (!targetTable) {
            console.warn(`Could not find table at index ${tableStartIndex}`);
            return null;
        }

        // Navigate to the specific cell
        const tableRows = targetTable.tableRows;
        if (!tableRows || rowIndex >= tableRows.length) {
            console.warn(`Row index ${rowIndex} out of bounds (table has ${tableRows?.length || 0} rows)`);
            return null;
        }

        const row = tableRows[rowIndex];
        const tableCells = row.tableCells;
        if (!tableCells || columnIndex >= tableCells.length) {
            console.warn(`Column index ${columnIndex} out of bounds (row has ${tableCells?.length || 0} columns)`);
            return null;
        }

        const cell = tableCells[columnIndex];
        const cellStartIndex = cell.startIndex;
        const cellEndIndex = cell.endIndex;

        // Get the content range within the cell
        let contentStartIndex = cellStartIndex;
        let contentEndIndex = cellEndIndex;
        let hasContent = false;

        if (cell.content && cell.content.length > 0) {
            const firstContent = cell.content[0];
            if (firstContent.startIndex !== undefined) {
                contentStartIndex = firstContent.startIndex;
            }

            // Check if there's actual text content
            for (const contentElement of cell.content) {
                if (contentElement.paragraph?.elements) {
                    for (const elem of contentElement.paragraph.elements) {
                        if (elem.textRun?.content && elem.textRun.content.trim()) {
                            hasContent = true;
                            break;
                        }
                    }
                }
            }

            // Get the last content element's end index
            const lastContent = cell.content[cell.content.length - 1];
            if (lastContent.endIndex !== undefined) {
                contentEndIndex = lastContent.endIndex;
            }
        }

        console.log(`Found cell (${rowIndex}, ${columnIndex}): cellStart=${cellStartIndex}, cellEnd=${cellEndIndex}, contentStart=${contentStartIndex}, contentEnd=${contentEndIndex}, hasContent=${hasContent}`);

        return {
            cellStartIndex,
            cellEndIndex,
            contentStartIndex,
            contentEndIndex,
            hasContent
        };
    } catch (error: any) {
        console.error(`Error getting table cell range: ${error.message}`);
        throw new Error(`Failed to get table cell range: ${error.message}`);
    }
}

/**
 * Edits a table cell's content and optionally applies styling
 * @param docs - Google Docs API client
 * @param documentId - The document ID
 * @param tableStartIndex - The starting index of the table
 * @param rowIndex - Row index (0-based)
 * @param columnIndex - Column index (0-based)
 * @param textContent - New text content for the cell (optional)
 * @param textStyle - Text styling to apply (optional)
 * @param paragraphStyle - Paragraph styling to apply (optional)
 * @returns Promise with the result
 */
export async function editTableCell(
    docs: Docs,
    documentId: string,
    tableStartIndex: number,
    rowIndex: number,
    columnIndex: number,
    textContent?: string,
    textStyle?: TextStyleArgs,
    paragraphStyle?: ParagraphStyleArgs
): Promise<{ success: boolean; message: string }> {
    // Get the cell range
    const cellRange = await getTableCellRange(docs, documentId, tableStartIndex, rowIndex, columnIndex);

    if (!cellRange) {
        throw new UserError(`Could not find cell at row ${rowIndex}, column ${columnIndex} in table at index ${tableStartIndex}`);
    }

    const requests: docs_v1.Schema$Request[] = [];

    // If we need to set text content
    if (textContent !== undefined) {
        // First, delete existing content if any (but preserve the paragraph structure)
        // We need to be careful here - cells always have at least one paragraph with a newline
        // We should delete content but not the structural newline

        if (cellRange.hasContent) {
            // Delete existing content (everything except the final newline)
            // The cell content typically ends with a newline we want to preserve
            const deleteEnd = cellRange.contentEndIndex - 1; // Keep the trailing newline
            if (deleteEnd > cellRange.contentStartIndex) {
                requests.push({
                    deleteContentRange: {
                        range: {
                            startIndex: cellRange.contentStartIndex,
                            endIndex: deleteEnd
                        }
                    }
                });
            }
        }

        // Insert new text at the cell's content start
        if (textContent) {
            requests.push({
                insertText: {
                    location: { index: cellRange.contentStartIndex },
                    text: textContent
                }
            });
        }
    }

    // Execute the batch update
    if (requests.length > 0) {
        await executeBatchUpdate(docs, documentId, requests);
    }

    // Now apply styles if requested (need to re-fetch document to get updated indices)
    if ((textStyle || paragraphStyle) && textContent) {
        // Re-fetch the cell range after content modification
        const updatedCellRange = await getTableCellRange(docs, documentId, tableStartIndex, rowIndex, columnIndex);

        if (updatedCellRange) {
            const styleRequests: docs_v1.Schema$Request[] = [];

            // Calculate the new text range
            const textStart = updatedCellRange.contentStartIndex;
            const textEnd = textStart + (textContent?.length || 0);

            // Apply text style
            if (textStyle && textContent) {
                const textStyleResult = buildUpdateTextStyleRequest(textStart, textEnd, textStyle);
                if (textStyleResult) {
                    styleRequests.push(textStyleResult.request);
                }
            }

            // Apply paragraph style
            if (paragraphStyle) {
                const paragraphStyleResult = buildUpdateParagraphStyleRequest(
                    updatedCellRange.contentStartIndex,
                    updatedCellRange.contentEndIndex,
                    paragraphStyle
                );
                if (paragraphStyleResult) {
                    styleRequests.push(paragraphStyleResult.request);
                }
            }

            if (styleRequests.length > 0) {
                await executeBatchUpdate(docs, documentId, styleRequests);
            }
        }
    }

    return {
        success: true,
        message: `Successfully edited cell (${rowIndex}, ${columnIndex}) in table at index ${tableStartIndex}`
    };
}

/**
 * Creates a table with data in one efficient operation
 * @param docs - Google Docs API client
 * @param documentId - The document ID
 * @param insertIndex - Where to insert the table
 * @param headers - Array of header strings
 * @param rows - 2D array of row data
 * @param boldHeaders - Whether to bold the header row (default true)
 * @param boldTotalRow - Whether to bold rows containing "Total" (default true)
 * @returns Promise with the result
 */
export async function createTableWithData(
    docs: Docs,
    documentId: string,
    insertIndex: number,
    headers: string[],
    rows: string[][],
    boldHeaders: boolean = true,
    boldTotalRow: boolean = true
): Promise<{ success: boolean; message: string; tableStartIndex: number }> {
    const numRows = rows.length + 1; // +1 for header
    const numCols = headers.length;

    if (numCols === 0) {
        throw new UserError("Headers array cannot be empty");
    }

    // Step 1: Insert the empty table
    const insertTableRequest: docs_v1.Schema$Request = {
        insertTable: {
            rows: numRows,
            columns: numCols,
            location: { index: insertIndex }
        }
    };
    await executeBatchUpdate(docs, documentId, [insertTableRequest]);

    // Step 2: Re-fetch document to get table structure
    const doc = await docs.documents.get({
        documentId,
    });

    if (!doc.data.body?.content) {
        throw new Error("Failed to fetch document content after table insertion");
    }

    // Step 3: Find the table we just inserted
    let tableElement: any = null;
    let tableStartIndex = 0;
    for (const element of doc.data.body.content) {
        if (element.table) {
            const startIdx = element.startIndex;
            if (startIdx != null && startIdx >= insertIndex - 5) {
                tableElement = element.table;
                tableStartIndex = startIdx;
                break;
            }
        }
    }

    if (!tableElement) {
        throw new Error("Could not find the inserted table");
    }

    // Step 4: Collect all cell data and their indices
    const cellInsertions: { index: number; text: string; isBold: boolean }[] = [];
    const tableRows = tableElement.tableRows || [];

    for (let rowIdx = 0; rowIdx < tableRows.length; rowIdx++) {
        const row = tableRows[rowIdx];
        const tableCells = row.tableCells || [];

        for (let colIdx = 0; colIdx < tableCells.length; colIdx++) {
            const cell = tableCells[colIdx];
            const cellContent = cell.content || [];

            if (cellContent.length > 0) {
                const firstContent = cellContent[0];
                // Get the correct index for text insertion - it's inside the paragraph element
                let cellStartIndex: number | null | undefined = null;
                if (firstContent.paragraph?.elements?.[0]?.startIndex != null) {
                    cellStartIndex = firstContent.paragraph.elements[0].startIndex;
                } else if (firstContent.startIndex != null) {
                    cellStartIndex = firstContent.startIndex;
                }

                // Determine the text to insert
                let cellText = '';
                let isBold = false;

                if (rowIdx === 0) {
                    // Header row
                    cellText = headers[colIdx] || '';
                    isBold = boldHeaders;
                } else {
                    // Data row
                    const dataRowIdx = rowIdx - 1;
                    if (dataRowIdx < rows.length && colIdx < rows[dataRowIdx].length) {
                        cellText = rows[dataRowIdx][colIdx] || '';
                    }
                    // Check if it's a total row
                    if (boldTotalRow && rows[dataRowIdx] && rows[dataRowIdx][0] &&
                        rows[dataRowIdx][0].toLowerCase().includes('total')) {
                        isBold = true;
                    }
                }

                if (cellText && cellStartIndex != null) {
                    cellInsertions.push({
                        index: cellStartIndex,
                        text: cellText,
                        isBold
                    });
                }
            }
        }
    }

    // Step 5: Sort by index descending (insert in reverse order to maintain indices)
    cellInsertions.sort((a, b) => b.index - a.index);

    // Step 6: Build insertText requests
    const textRequests: docs_v1.Schema$Request[] = cellInsertions.map(cell => ({
        insertText: {
            location: { index: cell.index },
            text: cell.text
        }
    }));

    // Execute text insertions
    if (textRequests.length > 0) {
        await executeBatchUpdate(docs, documentId, textRequests);
    }

    // Step 7: Re-fetch document and apply bold styling
    const boldCells = cellInsertions.filter(c => c.isBold);
    if (boldCells.length > 0) {
        // Re-fetch to get updated indices
        const docAfterText = await docs.documents.get({
            documentId,
        });

        if (docAfterText.data.body?.content) {
            const boldRequests: docs_v1.Schema$Request[] = [];

            for (const element of docAfterText.data.body.content) {
                if (element.table && element.startIndex != null &&
                    Math.abs(element.startIndex - tableStartIndex) <= 20) {
                    const tRows = element.table.tableRows || [];

                    for (let rowIdx = 0; rowIdx < tRows.length; rowIdx++) {
                        const row = tRows[rowIdx];
                        const cells = row.tableCells || [];

                        // Determine if this row should be bold
                        let shouldBoldRow = false;
                        if (rowIdx === 0 && boldHeaders) {
                            shouldBoldRow = true;
                        } else if (boldTotalRow && rowIdx > 0) {
                            const dataRowIdx = rowIdx - 1;
                            if (dataRowIdx < rows.length && rows[dataRowIdx][0] &&
                                rows[dataRowIdx][0].toLowerCase().includes('total')) {
                                shouldBoldRow = true;
                            }
                        }

                        if (shouldBoldRow) {
                            for (const cell of cells) {
                                const cellContent = cell.content || [];
                                for (const para of cellContent) {
                                    const elements = para.paragraph?.elements || [];
                                    for (const elem of elements) {
                                        if (elem.textRun?.content && elem.textRun.content.trim()) {
                                            const start = elem.startIndex;
                                            const end = elem.endIndex;
                                            if (start != null && end != null && end > start) {
                                                boldRequests.push({
                                                    updateTextStyle: {
                                                        range: { startIndex: start, endIndex: end },
                                                        textStyle: { bold: true },
                                                        fields: 'bold'
                                                    }
                                                });
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    break;
                }
            }

            if (boldRequests.length > 0) {
                await executeBatchUpdate(docs, documentId, boldRequests);
            }
        }
    }

    return {
        success: true,
        message: `Successfully created ${numRows}x${numCols} table with data at index ${insertIndex}`,
        tableStartIndex
    };
}
