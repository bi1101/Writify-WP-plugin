/**
 * Script Name: Docx Export for Speaking
 * Version: 1.0.3
 * Last Updated: 4-08-2024
 * Author: bi1101
 * Description: Export the result page as docx files with comments.
 */
var generatedBlob = null;
var globalCommentCounter = 0;
var transcriptText = null;

async function fetchStylesXML() {
    const response = await fetch('https://cdn.jsdelivr.net/gh/bi1101/Docx-Export-for-Writify/styles.xml');
    const xmlText = await response.text();
    return xmlText;
}

function createVocabSectionWithComments(rawComments) {
    const { Paragraph, TextRun, HeadingLevel, CommentRangeStart, CommentRangeEnd, CommentReference } = docx;

    // Get vocabulary score and prepare heading
    const vocabScore = document.querySelector('#vocab_score_wrap').innerText;
    const headingText = `Vocabulary: ${vocabScore}`;

    const sectionChildren = [
        new Paragraph({
            children: [new TextRun(headingText)],
            heading: HeadingLevel.HEADING_1
        })
    ];

    // Get transcript elements
    const transcriptWrap = document.querySelector('#vocab-transcript-wrap');
    const fileBlocks = transcriptWrap.querySelectorAll('.file-block');

    // Set to track processed errorIds
    const processedErrorIds = new Set();

    fileBlocks.forEach(fileblock => {
        // Add file block heading (file title)
        const fileTitle = fileblock.querySelector('.file-title').innerText;
        sectionChildren.push(
            new Paragraph({
                children: [new TextRun({ text: fileTitle, bold: true })]
            })
        );

        const paraChildren = [];
        // Process each transcript in the file block
        const transcripts = fileblock.querySelectorAll('p.transcript-text');
        transcripts.forEach((transcriptEl, index) => {
            let transcriptText = transcriptEl.textContent;
            if (index === 0) {
                transcriptText = transcriptEl.textContent.trim();
            }
            const errors = transcriptEl.querySelectorAll('span');
            let currentPosition = 0;

            if (errors.length > 0) {
                // Process each error in the transcript
                errors.forEach(error => {
                    const errorText = error.textContent;
                    const errorID = error.id;
                    const suggestionID = errorID.replace('ERROR_', 'VOCAB_ERROR_');

                    // Check if the errorId has already been processed
                    if (!processedErrorIds.has(suggestionID)) {
                        // Mark this errorId as processed
                        processedErrorIds.add(suggestionID);

                        // Find corresponding comment from rawComments
                        const currentErrorComment = rawComments.find(comment => comment.errorId === suggestionID);

                        if (currentErrorComment) {
                            const wordCount = errorText.trim().split(/\s+/).length;

                            // Escape special characters in errorText
                            const escapedErrorText = errorText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

                            // Build regex for word boundary or non-boundary search
                            let regex;
                            if (wordCount <= 10) {
                                // Word boundary-based search for 10 words or fewer
                                regex = new RegExp(`\\b${escapedErrorText}\\b(\\s*[.,;?!]*)`, 'i');
                            } else {
                                // Non-word boundary search for more than 10 words
                                regex = new RegExp(`${escapedErrorText}(\\s*[.,;?!]*)`, 'i');
                            }

                            // Slice transcriptText from currentPosition
                            const transcriptSlice = transcriptText.slice(currentPosition);

                            // Search for the errorText in the sliced transcript
                            const match = transcriptSlice.match(regex);

                            if (match) {
                                // Calculate the position in the original transcriptText
                                const commentStartPos = currentPosition + match.index;

                                if (commentStartPos >= 0) {
                                    // Add the text before the comment
                                    if (commentStartPos > currentPosition) {
                                        paraChildren.push(new TextRun(transcriptText.slice(currentPosition, commentStartPos)));
                                    }

                                    // Add the comment with the error
                                    paraChildren.push(new CommentRangeStart(currentErrorComment.commentCounter));
                                    paraChildren.push(new TextRun(currentErrorComment.orignal));
                                    paraChildren.push(new CommentRangeEnd(currentErrorComment.commentCounter));
                                    paraChildren.push(new TextRun({
                                        children: [new CommentReference(currentErrorComment.commentCounter)]
                                    }));

                                    // Update current position after the comment
                                    currentPosition = commentStartPos + currentErrorComment.orignal.length;
                                }
                            }
                        }
                    }
                });

                // Add any remaining text after the last comment
                if (currentPosition < transcriptText.length) {
                    paraChildren.push(new TextRun(transcriptText.slice(currentPosition)));
                }
            } else {
                // If no errors, add the whole transcript text
                paraChildren.push(new TextRun(transcriptText));
            }
        });

        // Push the paragraph with content and comments
        sectionChildren.push(new Paragraph({
            children: paraChildren,
            indent: {
                firstLine: 0, // Set first line indent to 0 to remove any indentation
            }
        }));
    });

    return sectionChildren;
}

function createGrammerSectionWithComments(rawComments) {
    const { Paragraph, TextRun, HeadingLevel, CommentRangeStart, CommentRangeEnd, CommentReference } = docx;

    // Get grammar score and prepare heading
    const grammerScore = document.querySelector('#grammer_score_wrap').innerText;
    const headingText = `Grammar: ${grammerScore}`;

    const sectionChildren = [
        new Paragraph({
            children: [new TextRun(headingText)],
            heading: HeadingLevel.HEADING_1
        })
    ];

    // Get transcript elements
    const transcriptWrap = document.querySelector('#grammer-transcript-wrap');
    const fileBlocks = transcriptWrap.querySelectorAll('.file-block');

    fileBlocks.forEach(fileblock => {
        // Add file block heading (file title)
        const fileTitle = fileblock.querySelector('.file-title').innerText;
        sectionChildren.push(
            new Paragraph({
                children: [new TextRun({ text: fileTitle, bold: true })]
            })
        );

        const paraChildren = [];
        // Process each transcript in the file block
        const transcripts = fileblock.querySelectorAll('p.transcript-text');
        transcripts.forEach((transcriptEl, index) => {
            let transcriptText = transcriptEl.textContent;
            if (index === 0) {
                transcriptText = transcriptEl.textContent.trim();
            }
            const errors = transcriptEl.querySelectorAll('span');
            let currentPosition = 0;

            if (errors.length > 0) {
                // Process each error in the transcript
                errors.forEach(error => {
                    const errorText = error.textContent;
                    const errorID = error.id;
                    const suggestionID = errorID.replace('ERROR_', 'GRAMMER_ERROR_');

                    // Find corresponding comment from rawComments
                    const currentErrorComment = rawComments.find(comment => comment.errorId === suggestionID);
                    if (currentErrorComment) {
                        const wordCount = errorText.trim().split(/\s+/).length;

                        // Escape special characters in errorText for regex
                        const escapedErrorText = errorText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

                        // Build regex based on word count
                        let regex;
                        if (wordCount <= 10) {
                            // Word boundary-based search for 10 words or fewer
                            regex = new RegExp(`\\b${escapedErrorText}\\b(\\s*[.,;?!]*)`, 'i');
                        } else {
                            // Non-word boundary search for more than 10 words
                            regex = new RegExp(`${escapedErrorText}(\\s*[.,;?!]*)`, 'i');
                        }

                        // Slice transcriptText from currentPosition and search for errorText
                        const transcriptSlice = transcriptText.slice(currentPosition);
                        const match = transcriptSlice.match(regex);

                        if (match) {
                            // Calculate the position in the original transcriptText
                            const commentStartPos = currentPosition + match.index;

                            if (commentStartPos >= 0) {
                                // Add the text before the comment
                                if (commentStartPos > currentPosition) {
                                    paraChildren.push(new TextRun(transcriptText.slice(currentPosition, commentStartPos)));
                                }

                                // Add the comment with the error
                                paraChildren.push(new CommentRangeStart(currentErrorComment.commentCounter));
                                paraChildren.push(new TextRun(currentErrorComment.orignal));
                                paraChildren.push(new CommentRangeEnd(currentErrorComment.commentCounter));
                                paraChildren.push(new TextRun({
                                    children: [new CommentReference(currentErrorComment.commentCounter)]
                                }));

                                // Update current position after the comment
                                currentPosition = commentStartPos + currentErrorComment.orignal.length;
                            }
                        }
                    }
                });

                // Add any remaining text after the last comment
                if (currentPosition < transcriptText.length) {
                    paraChildren.push(new TextRun(transcriptText.slice(currentPosition)));
                }
            } else {
                // If no errors, add the whole transcript text
                paraChildren.push(new TextRun(transcriptText));
            }
        });

        // Push the paragraph with content and comments
        sectionChildren.push(new Paragraph({
            children: paraChildren,
            indent: {
                firstLine: 0, // Set first line indent to 0 to remove any indentation
            }
        }));
    });

    return sectionChildren;
}


function createPronunSectionWithComments(rawComments) {
    const { Paragraph, TextRun, HeadingLevel, CommentRangeStart, CommentRangeEnd, CommentReference } = docx;

    console.log("Starting Pronunciation Section Creation...");

    // Get pronunciation score and prepare heading
    const pronunScore = document.querySelector('#pronun_score_wrap').innerText;
    const headingText = `Pronunciation: ${pronunScore}`;
    console.log(`Pronunciation score heading added: ${headingText}`);

    const sectionChildren = [
        new Paragraph({
            children: [new TextRun(headingText)],
            heading: HeadingLevel.HEADING_1
        })
    ];

    // Get transcript elements
    const transcriptWrap = document.querySelector('#pronun-transcript-wrap');
    const fileBlocks = transcriptWrap.querySelectorAll('.file-block');

    // Set to track processed errorIds
    const processedErrorIds = new Set();

    fileBlocks.forEach(fileblock => {
        // Add file block heading (file title)
        const fileTitle = fileblock.querySelector('.file-title').innerText;
        sectionChildren.push(
            new Paragraph({
                children: [new TextRun({ text: fileTitle, bold: true })]
            })
        );
        console.log(`Processing file block: ${fileTitle}`);

        const paraChildren = [];
        // Process each transcript in the file block
        const transcripts = fileblock.querySelectorAll('p.transcript-text');
        transcripts.forEach(transcriptEl => {
            const transcriptText = getTextContent(transcriptEl); // Get complete transcript content
            console.log(`Transcript content: ${transcriptText}`);
            const errors = transcriptEl.querySelectorAll('span.pronun-error');
            console.log(`Found ${errors.length} pronunciation errors in the transcript.`);
            let currentPosition = 0;

            if (errors.length > 0) {
                // Process each error in the transcript
                errors.forEach(error => {
                    const errorText = error.textContent;
                    const errorID = error.id;
                    const suggestionID = errorID.replace('PRONUN_ERROR_', 'SIDEPANEL_PRONUN_ERROR_');
                    const errorPosition = parseInt(error.getAttribute('position')); // Get the stored position
                    console.log(`Processing error with suggestionID: ${suggestionID}, text: "${errorText}", position: ${errorPosition}`);

                    // Check if the errorId has already been processed
                    if (!processedErrorIds.has(suggestionID)) {
                        console.log(`Processing new error with suggestionID: ${suggestionID}`);
                        // Mark this errorId as processed
                        processedErrorIds.add(suggestionID);

                        // Find corresponding comment from rawComments
                        const currentErrorComment = rawComments.find(comment => comment.errorId === suggestionID);
                        if (currentErrorComment) {
                            console.log(`Found matching comment for suggestionID: ${suggestionID}`);

                            // Handle any text before the comment
                            if (errorPosition > currentPosition) {
                                paraChildren.push(new TextRun(transcriptText.slice(currentPosition, errorPosition)));
                                console.log(`Added text before comment: ${transcriptText.slice(currentPosition, errorPosition)}`);
                            }

                            // Add the comment with the error
                            paraChildren.push(new CommentRangeStart(currentErrorComment.commentCounter));
                            paraChildren.push(new TextRun(`${currentErrorComment.orignal}`));
                            paraChildren.push(new CommentRangeEnd(currentErrorComment.commentCounter));
                            paraChildren.push(new TextRun({
                                children: [new CommentReference(currentErrorComment.commentCounter)]
                            }));
                            console.log(`Added comment for error: ${currentErrorComment.orignal}`);

                            // Update current position after the comment
                            currentPosition = errorPosition + currentErrorComment.orignal.length;
                            console.log(`Updated current position to: ${currentPosition}`);
                        } else {
                            console.log(`No matching comment found for suggestionID: ${suggestionID}`);
                        }
                    } else {
                        console.log(`Skipping already processed suggestionID: ${suggestionID}`);
                    }
                });

                // Add any remaining text after the last comment
                if (currentPosition < transcriptText.length) {
                    paraChildren.push(new TextRun(transcriptText.slice(currentPosition)));
                    console.log(`Added remaining text after last comment: ${transcriptText.slice(currentPosition)}`);
                }
            } else {
                // If no errors, add the whole transcript text
                paraChildren.push(new TextRun(transcriptText));
                console.log(`No errors found. Added full transcript text.`);
            }
        });

        // Push the paragraph with content and comments
        sectionChildren.push(new Paragraph({
            children: paraChildren,
            indent: {
                firstLine: 0, // Set first line indent to 0 to remove any indentation
            }
        }));
        console.log(`Completed processing of file block: ${fileTitle}`);
    });

    console.log("Completed pronunciation section creation.");
    return sectionChildren;
}

function createNormalSections(elementId) {
    const element = document.getElementById(elementId);
    if (!element) {
        console.warn(`No element found with ID: ${elementId}`);
        return [];
    }

    const sections = [];
    element.childNodes.forEach((child) => {
        if (child.nodeType === 1) {
            // Check if the node is an element
            if (child.tagName === "P") {
                // For paragraph tags
                sections.push(htmlParagraphToDocx(child.outerHTML));
            } else if (child.tagName === "OL" || child.tagName === "UL") {
                // For ordered or unordered lists
                sections.push(...bulletPointsToDocx(child.outerHTML));
            } else if (child.tagName === "TABLE") {
                // For table tags
                sections.push(convertTableToDocx(child));
            }
        }
    });

    return sections;
}

function convertTableToDocx(tableElement) {
    if (!tableElement) {
        return null;
    }

    const rows = [];

    // Helper function to create a TableCell
    const createTableCell = (text, isHeader = false) => {
        return new docx.TableCell({
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: text,
                            bold: isHeader, // Bold text for header cells
                        })
                    ]
                })
            ],
        });
    };

    // Helper function to process rows
    const processRows = (rowElements, isHeader = false) => {
        return Array.from(rowElements).map((row) => {
            const cells = Array.from(row.querySelectorAll("th, td")).map((cell) => {
                const cellText = cell?.innerText?.trim() || "N/A";
                return createTableCell(cellText, isHeader);
            });
            return new docx.TableRow({
                children: cells,
                tableHeader: isHeader,
                cantSplit: true // Prevents row from splitting across pages
            });
        });
    };

    // Process thead rows (header rows)
    const thead = tableElement.querySelector("thead");
    if (thead) {
        const headerRows = processRows(thead.querySelectorAll("tr"), true);
        rows.push(...headerRows);
    }

    // Process tbody rows (body rows)
    const tbody = tableElement.querySelector("tbody") || tableElement;
    const bodyRows = processRows(tbody.querySelectorAll("tr"));
    rows.push(...bodyRows);

    // Create and return the table
    return new docx.Table({
        rows,
        width: { size: 100, type: docx.WidthType.PERCENTAGE }
    });
}

function htmlParagraphToDocx(htmlContent) {
    // Convert the HTML string into a DOM element
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = htmlContent;

    const paragraph = tempDiv.querySelector("p");
    if (!paragraph) {
        console.warn("No paragraph element found in the provided HTML content.");
        return;
    }

    // Use processNodeForFormatting to handle child nodes
    const children = [];
    Array.from(paragraph.childNodes).forEach((child) => {
        children.push(...processNodeForFormatting(child));
    });

    return new docx.Paragraph({ children });
}

function processNodeForFormatting(node) {
    let textRuns = [];

    // Handle text nodes
    if (node.nodeType === 3) {
        // Node type 3 is a Text node
        textRuns.push(new docx.TextRun(node.nodeValue));
    }

    // Handle element nodes like <strong>, <em>, etc.
    else if (node.nodeType === 1) {
        // Node type 1 is an Element node
        const textContent = node.innerText;

        // Basic formatting options
        let formattingOptions = {};

        // Check the tag to determine formatting
        switch (node.tagName) {
            case "STRONG":
            case "B":
                formattingOptions.bold = true;
                break;
            case "EM":
            case "I":
                formattingOptions.italic = true;
                break;
            case "U":
                formattingOptions.underline = {
                    color: "auto",
                    type: docx.UnderlineType.SINGLE
                };
                break;
            // Add cases for other formatting tags as needed
        }

        // Check for nested formatting
        if (node.children.length > 0) {
            Array.from(node.childNodes).forEach((childNode) => {
                textRuns.push(...processNodeForFormatting(childNode));
            });
        } else {
            textRuns.push(
                new docx.TextRun({
                    text: textContent,
                    ...formattingOptions
                })
            );
        }
    }

    return textRuns;
}

function processList(list, level, paragraphs) {
    Array.from(list.children).forEach((item) => {
        paragraphs.push(createBulletPointParagraphs(item, level));

        // Process nested lists
        const nestedList = item.querySelector("ul, ol");
        if (nestedList) {
            processList(nestedList, level + 1, paragraphs);
        }
    });
}

function createBulletPointParagraphs(item, level) {
    let contentTextRuns = [];

    // Check if the item contains a paragraph element
    const paragraphElement = item.querySelector("p");

    if (paragraphElement) {
        contentTextRuns = processNodeForFormatting(paragraphElement);
    } else {
        Array.from(item.childNodes).forEach((childNode) => {
            contentTextRuns.push(...processNodeForFormatting(childNode));
        });
    }

    return new docx.Paragraph({
        children: contentTextRuns,
        bullet: {
            level: level
        }
    });
}

function bulletPointsToDocx(outerHTML) {
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = outerHTML;

    const docxItems = [];

    // Check whether the provided outerHTML is an ordered or unordered list and process it accordingly
    const listElement = tempDiv.querySelector("ol, ul");
    if (listElement) {
        processList(listElement, 0, docxItems);
    } else {
        console.warn(
            "Provided HTML does not contain a valid list element (ol or ul)."
        );
        return [];
    }

    return docxItems; // Ensure we return the docxItems
}

function createHeaderParagraph(text) {
    const { Paragraph, TextRun, HeadingLevel } = docx;
    return new Paragraph({
        children: [new TextRun(text)],
        heading: HeadingLevel.HEADING_1
    });
}

function getTextContent(element) {
    let textWithSpaces = '';

    element.childNodes.forEach(node => {
        if (node.nodeType === Node.TEXT_NODE) {
            // Append text nodes directly
            textWithSpaces += node.textContent;
        } else if (node.nodeType === Node.ELEMENT_NODE && node.tagName === 'SPAN') {
            // Append span text and add a space after each span
            textWithSpaces += node.textContent + ' ';
        }
    });

    // Trim to remove any trailing spaces
    return textWithSpaces.trim();
}

function createFluencySection() {
    const { Paragraph, TextRun, HeadingLevel } = docx;
    let score = document.querySelector('#fluency_score_wrap').innerText;
    let headingText = `Fluency: ${score}`;
    let wpm = document.querySelector('#wpmValue span').textContent;
    let wpmData = wpm.match(/(\d+)\s*WPM\s*(.*)/);
    wpm = wpmData[1];
    let speedText = wpmData[2];
    let wpmColor;
    let fillerWordsColor = '#ffc000';
    let badPause = new TextRun({
        text: '_',
        color: '#FF4949'
    });
    let missedPause = new TextRun({
        text: '_',
        color: '#ffc000'
    });
    if (parseInt(wpm) < 120) {
        // Red if less then 120
        wpmColor = '#FF4949';
    } else if (parseInt(wpm) > 180) {
        // Orange if greater then 180
        wpmColor = '#FF822E';
    } else {
        // Green if between 120 and 180
        wpmColor = '#13CE66';
    }
    let sectionChildren = [
        new Paragraph({
            children: [new TextRun(headingText)],
            heading: HeadingLevel.HEADING_1
        }),
        new Paragraph({
            children: [
                new TextRun('Speaking rate:'),
                new TextRun({
                    text: wpm,
                    color: wpmColor
                })
            ]

        }),
        new Paragraph({
            children: [
                new TextRun('Your Pace Was '),
                new TextRun({
                    text: speedText,
                    color: wpmColor
                }),
                new TextRun(' during this recording'),
            ]
        }),
        new Paragraph({
            alignment: docx.AlignmentType.RIGHT,
            children: [
                new TextRun('Note: '),
                new TextRun({
                    text: 'Uhm, uh',
                    color: fillerWordsColor
                }),
                new TextRun('      '),
                badPause,
                new TextRun('Bad Pause'),
                new TextRun('      '),
                missedPause,
                new TextRun('Missed Pause'),
            ]
        }),
    ];

    // Get transcript elements
    const transcriptWrap = document.querySelector('#fluency-transcript-wrap');
    const fileBlocks = transcriptWrap.querySelectorAll('.file-block');

    fileBlocks.forEach(fileblock => {
        // Add file block heading (file title)
        const fileTitle = fileblock.querySelector('.file-title').innerText;
        sectionChildren.push(
            new Paragraph({
                children: [new TextRun({ text: fileTitle, bold: true })]
            })
        );
        const paraChildren = [];
        // Process each transcript in the file block
        const transcripts = fileblock.querySelectorAll('p.transcript-text');
        transcripts.forEach(transcriptEl => {
            let transcriptText = transcriptEl.textContent;
            transcriptText = transcriptText.replace(/\s+/g, ' ').trim();
            const errorSpans = transcriptEl.querySelectorAll('span');
            let currentPosition = 0;
            if (errorSpans.length > 0) {
                // Process each error in the transcript
                errorSpans.forEach((error, index) => {
                    let isPauseError = false;
                    let isBadPause = false;
                    let isMissingPause = false;
                    if (error.classList.contains('pause-error')) {
                        isPauseError = true;
                        if (error.classList.contains('unexpectedbreak-pause')) {
                            isBadPause = true;
                        } else {
                            isMissingPause = true;
                        }
                    }
                    let errorText = '';
                    if (isPauseError) {
                        // Get the index of the pause error span in the transcript text
                        const textWithPlaceholders = getTextWithPlaceholders(transcriptEl);
                        let placeholder = `<<SPAN_${index}>>`;
                        let pauseErrorStartPos = textWithPlaceholders.indexOf(placeholder, currentPosition);
                        if (index > 0) {
                            pauseErrorStartPos = pauseErrorStartPos - (placeholder.length + 1); // +1 for Space
                        }
                        console.log(`Current Position: ${currentPosition} Error Position :${pauseErrorStartPos} in ${textWithPlaceholders}`);
                        if (pauseErrorStartPos >= 0) {
                            // Add the text before the comment
                            if (pauseErrorStartPos > currentPosition) {
                                let slicedText = transcriptText.slice(currentPosition, pauseErrorStartPos);
                                console.log(slicedText);
                                paraChildren.push(new TextRun(slicedText));
                            }
                            if (isMissingPause) {
                                paraChildren.push(missedPause);
                            } else if (isBadPause) {
                                paraChildren.push(badPause);
                            }
                            // Update current position
                            currentPosition = pauseErrorStartPos;
                        }
                    } else {
                        let errorText = error.innerText;
                        console.log(errorText);
                        const fillerWordStartPos = transcriptText.toLowerCase().indexOf(errorText.toLowerCase(), currentPosition);
                        if (fillerWordStartPos >= 0) {
                            // Add the text before the comment
                            if (fillerWordStartPos > currentPosition) {
                                paraChildren.push(new TextRun(transcriptText.slice(currentPosition, fillerWordStartPos)));
                            }
                            paraChildren.push(new TextRun({
                                text: `${errorText}`,
                                color: '#ffc000'
                            }));
                            // Update current position after the comment
                            currentPosition = fillerWordStartPos + errorText.length;
                        }
                    }
                });
                // Add any remaining text after the last comment
                if (currentPosition < transcriptText.length) {
                    paraChildren.push(new TextRun(transcriptText.slice(currentPosition)));
                }
            } else {
                // If no errors, add the whole transcript text
                paraChildren.push(new TextRun(transcriptText));
            }
        });
        // Push the paragraph with content and comments
        sectionChildren.push(new Paragraph({
            children: paraChildren,
            indent: {
                firstLine: 0, // Set first line indent to 0 to remove any indentation
            }
        }));
    });
    return sectionChildren;
}

function getTranscriptText() {
    let text = '';
    let fileBlocks = document.querySelectorAll('#vocab-transcript-wrap .file-block');
    fileBlocks.forEach(fileBlock => {
        let transcriptEntries = fileBlock.querySelectorAll('.transcript-entry .transcript-text');
        transcriptEntries.forEach(transcriptHTML => {
            text += transcriptHTML.textContent + '\n';
        });

        text += '\n\n'; // Double Line Break After File
    });
    text = text.toLowerCase();
    transcriptText = text;
    return text;
}

function getRawVocabComments() {
    const suggestions = document.querySelectorAll('#vocab-suggestions .upgrade_vocab');
    let transcript = transcriptText ?? getTranscriptText();
    let comments = [];
    const seenErrorIds = new Set(); // Track seen errorIds to ensure uniqueness

    suggestions.forEach(suggestionEl => {
        let orignal = suggestionEl.querySelector('.original-vocab').innerText;
        let improved = suggestionEl.querySelector('.improved-vocab').innerText;
        let explanation = suggestionEl.querySelector('.explanation').innerText;

        // Get the errorId
        const errorId = suggestionEl.id;

        // Find the position of the original vocabulary in the essay text
        const position = transcript.indexOf(orignal.toLowerCase());

        // Only add the comment if the original vocabulary exists in the transcript
        // and the errorId hasn't been added before (for uniqueness)
        if (position >= 0 && !seenErrorIds.has(errorId)) {
            let commentObj = {
                // Specific to Vocab
                orignal: orignal,
                improved: improved,
                explanation: explanation,
                // Common Properties
                errorId: errorId,
                commentType: 'vocab',
                commentedText: orignal,
                commentContent: `${orignal} -> ${improved} \n Explanation: ${explanation}`,
                commentCounter: globalCommentCounter,
                position: position // Store Position
            };

            // Mark this errorId as seen
            seenErrorIds.add(errorId);

            globalCommentCounter++;
            comments.push(commentObj);
        }
    });

    // Sort comments based on their position
    comments.sort((a, b) => a.position - b.position);

    return comments;
}

function getRawGrammerComments() {
    const suggestions = document.querySelectorAll('#grammer-suggestions .grammer_error');
    let transcript = transcriptText ?? getTranscriptText();
    let comments = [];
    const seenErrorIds = new Set(); // Track seen errorIds to ensure uniqueness

    suggestions.forEach(suggestionEl => {
        let orignal = suggestionEl.querySelector('.original-grammer-word').innerText;
        let improved = '';
        let improvedWords = suggestionEl.querySelectorAll('.improved-grammer-word span');
        improvedWords.forEach((improvedWord, index) => {
            improved += ` ${improvedWord.innerText}`;
            if (index != (improvedWords.length - 1)) {
                improved += ',';
            }
        });
        let explanation = suggestionEl.querySelector('.explanation').innerText;

        // Get the errorId
        const errorId = suggestionEl.id;

        // Find the position of the original grammar error in the transcript
        const position = transcript.indexOf(orignal.toLowerCase());

        // Only add the comment if the original grammar error exists in the transcript
        // and the errorId hasn't been added before (for uniqueness)
        if (position >= 0 && !seenErrorIds.has(errorId)) {
            let commentObj = {
                // Specific to Grammar
                orignal: orignal,
                improved: improved.trim(),
                explanation: explanation,
                // Common Properties
                errorId: errorId,
                commentType: 'grammer',
                commentedText: orignal,
                commentContent: `${orignal} -> ${improved} \n Explanation: ${explanation}`,
                commentCounter: globalCommentCounter,
                position: position // Store Position
            };

            // Mark this errorId as seen
            seenErrorIds.add(errorId);

            globalCommentCounter++;
            comments.push(commentObj);
        }
    });

    // Sort comments based on their position in the transcript
    comments.sort((a, b) => a.position - b.position);

    return comments;
}

function getTextWithPlaceholders(element) {
    // Clone the element to avoid modifying the original HTML content
    const clonedElement = element.cloneNode(true);

    // Replace all span tags with a placeholder
    clonedElement.querySelectorAll('span').forEach((span, index) => {
        const placeholder = `<<SPAN_${index}>>`;
        span.replaceWith(placeholder);
    });

    // Return the text content with placeholders
    return clonedElement.textContent.replace(/\s+/g, ' ').trim();;
}

// Function to get the start index of the placeholder
function getSpanStartIndexInText(transcriptEl, span) {
    // Replace spans with placeholders and get the resulting text content
    const textWithPlaceholders = getTextWithPlaceholders(transcriptEl);
    console.log(textWithPlaceholders);
    // Generate the placeholder for the specific span
    const allSpans = Array.from(transcriptEl.querySelectorAll('span'));
    const spanIndex = allSpans.indexOf(span); // Get the index of the specific span
    const placeholder = `<<SPAN_${spanIndex}>>`;

    // Get the index of the placeholder in the text content
    const placeholderStartPos = textWithPlaceholders.indexOf(placeholder);

    return placeholderStartPos;
}

function getRawPronunComments() {
    const suggestions = document.querySelectorAll('#pronun-suggestions .pronun_error');
    let transcript = transcriptText ?? getTranscriptText();
    let comments = [];
    const seenErrorIds = new Set(); // Track seen errorIds to ensure uniqueness

    suggestions.forEach(suggestionEl => {
        let orignal = suggestionEl.querySelector('.sidepanel-orignal-pronun').getAttribute('data-text');
        let improved = suggestionEl.querySelector('.sidepanel-correct-pronun').innerText;
        let explanation = ''; // Explanation is empty for pronunciation

        // Get the errorId
        const errorId = suggestionEl.id;

        // Find the position of the original pronunciation error in the transcript
        const position = transcript.indexOf(orignal.toLowerCase());

        // Only add the comment if the original pronunciation error exists in the transcript
        // and the errorId hasn't been added before (for uniqueness)
        if (position >= 0 && !seenErrorIds.has(errorId)) {
            let commentObj = {
                // Specific to Pronunciation
                orignal: orignal,
                improved: improved,
                explanation: explanation,
                // Common Properties
                errorId: errorId,
                commentType: 'pronunciation',
                commentedText: orignal,
                commentContent: `${orignal} -> ${improved}`,
                commentCounter: globalCommentCounter,
                position: position // Store Position
            };

            // Mark this errorId as seen
            seenErrorIds.add(errorId);

            globalCommentCounter++;
            comments.push(commentObj);
        }
    });

    // Sort comments based on their position in the transcript
    comments.sort((a, b) => a.position - b.position);

    return comments;
}



function convertVocabCommentstoDocxFormat(rawComments) {
    // Use the passed user data
    const authorName = writifyUserData.firstName + ' ' + writifyUserData.lastName || "IELTS Science";
    return rawComments.map((comment, index) => ({
        id: comment.commentCounter,
        author: authorName.trim() ?? 'IELTS Science', // Replaced "Teacher" with the user's name
        date: new Date(),
        children: [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: `${comment.orignal} -> ${comment.improved}`
                    })
                ]
            }),
            new docx.Paragraph({}),
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: `${comment.explanation}`
                    })
                ]
            }),
        ]
    }));
}

function convertGrammerCommentstoDocxFormat(rawComments) {
    // Use the passed user data
    const authorName = writifyUserData.firstName + ' ' + writifyUserData.lastName || "IELTS Science";
    return rawComments.map((comment, index) => ({
        id: comment.commentCounter,
        author: authorName.trim() ?? 'IELTS Science', // Replaced "Teacher" with the user's name
        date: new Date(),
        children: [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: `${comment.orignal} -> ${comment.improved}`
                    })
                ]
            }),
            new docx.Paragraph({}),
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: `${comment.explanation}`
                    })
                ]
            }),
        ]
    }));
}

function convertPronunCommentstoDocxFormat(rawComments) {
    // Use the passed user data
    const authorName = writifyUserData.firstName + ' ' + writifyUserData.lastName || "IELTS Science";
    return rawComments.map((comment, index) => ({
        id: comment.commentCounter,
        author: authorName.trim() ?? 'IELTS Science', // Replaced "Teacher" with the user's name
        date: new Date(),
        children: [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: `Pronunciation Error:`
                    })
                ]
            }),
            new docx.Paragraph({}),
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: `${comment.orignal} -> ${comment.improved}`
                    })
                ]
            }),
        ]
    }));
}



async function exportDocument(saveBlob = true) {
    const customStyles = await fetchStylesXML();

    // Vocab comments
    let vocabRawComments = getRawVocabComments();
    let vocabComments = convertVocabCommentstoDocxFormat(vocabRawComments);
    // Grammer Comments
    let grammerRawComments = getRawGrammerComments();
    let grammerComments = convertGrammerCommentstoDocxFormat(grammerRawComments);
    // Pronun Comments
    let pronunRawComments = getRawPronunComments();
    let pronunComments = convertPronunCommentstoDocxFormat(pronunRawComments);


    let rawComments = [...vocabRawComments, ...grammerRawComments, ...pronunRawComments];
    console.log(pronunRawComments);
    let formattedComments = [...vocabComments, ...grammerComments, ...pronunComments];
    let sectionChildren = [
        ...createVocabSectionWithComments(vocabRawComments),
        ...createGrammerSectionWithComments(grammerRawComments),
        ...createPronunSectionWithComments(pronunRawComments),
        ...createFluencySection(),
    ];

    sectionChildren.push(createHeaderParagraph('Improved Answer'));
    sectionChildren.push(...createNormalSections('improved-ans-wrap'));

    // console.log(rawComments);
    // console.log(formattedComments);
    const doc = new docx.Document({
        title: "Result", // Adjust as needed
        externalStyles: customStyles, // Use externalStyles instead of styles
        comments: {
            children: formattedComments
        },

        sections: [
            {
                properties: {},
                children: sectionChildren
            }
        ]
    });

    // Convert the document to a blob and save it
    generatedBlob = await docx.Packer.toBlob(doc);
    if (saveBlob == true) {
        saveBlobAsDocx(generatedBlob);
    }
}

async function getGeneratedBlob() {
    if (!generatedBlob) {
        await exportDocument(false);
    }
    return generatedBlob;
}

function createHeading(text) {
    const { Paragraph, TextRun, HeadingLevel } = docx;
    return new Paragraph({
        children: [new TextRun(text)],
        heading: HeadingLevel.HEADING_1
    });
}

function saveBlobAsDocx(blob) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Result.docx";
    document.body.appendChild(a);
    a.click();
    a.remove();
}

// Function to Export Document

function startDocumentExport() {
    console.log('Document Export Started');
    exportDocument().catch(error => console.error(error));
}
