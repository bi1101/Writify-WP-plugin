/**
 * Script Name: Grammer Interaction Handler
 * Version: 1.0.0
 * Last Updated: 13-11-2023
 * Author: bi1101
 * Description: Turn static output from languagetool into intereactive clickable div.
 */

// Store frequently used selectors
const $grammerDocument = jQuery(document);
const $grammertranscriptWrap = jQuery('#grammer-transcript-wrap');
const $grammerErrorsWrap = jQuery('#grammer-suggestions');

function generateGrammarErrorHTML(matches) {
    let errorCounter = 1;
    let html = '';
    let notFoundinTranscriptErrorIDs = [];
    if(matches.length < 1){
        jQuery('#grammer-suggestions img').replaceWith('<span>No Errors Found</span>');
        return;
    }
    
    function wrapErrorWordInTranscript(word, context, transcript, errorId) {
        const highlightedWord = `<span class="grammer-error" id="ERROR_${errorId}">${word}</span>`;
        const matchThreshold = 0.7; // Set threshold for minimum match (70% similarity)
    
        // Regex to find if the word is already inside a <span>
        const spanRegex = new RegExp(`<span[^>]*?>[^<]*?\\b${word}\\b[^<]*?</span>`, 'gi');
    
        // Function to normalize the context and transcript for better comparison
        function cleanString(str) {
            return str.replace(/[\.\,\!\?\;\:]/g, '') // Remove punctuation
                      .replace(/\s+/g, ' ') // Normalize spaces
                      .trim().toLowerCase(); // Trim and convert to lowercase
        }
    
        // Function to compute similarity percentage between two strings
        function getMatchPercentage(str1, str2) {
            let matchedWords = 0;
            const words1 = str1.split(' ');
            const words2 = str2.split(' ');
    
            words1.forEach(word => {
                if (words2.includes(word)) matchedWords++;
            });
    
            return matchedWords / Math.max(words1.length, words2.length);
        }
    
        // Clean context and transcript
        const cleanContext = cleanString(context);
        const cleanTranscript = cleanString(transcript);
    
        // Handle the rare case: If the word/phrase is more than 30 words long, skip context matching
        if (word.split(' ').length > 10) {
            let wordRegex = new RegExp(`${word}`, 'gi'); // Exact word match regex
            let match;
    
            while ((match = wordRegex.exec(transcript)) !== null) {
                let wordStartIndex = match.index;
                let wordEndIndex = wordStartIndex + word.length;
    
                // Skip if word is already inside a span tag
                let precedingText = transcript.slice(0, wordStartIndex);
                let followingText = transcript.slice(wordEndIndex);
    
                if (spanRegex.test(precedingText + word + followingText)) {
                    continue;
                }
    
                // Log when the word is found and wrapped
                console.log({
                    word: word,
                    context: context,
                    cleanedContext: cleanContext,
                    transcript: transcript,
                    errorId: errorId,
                    wordPosition: {
                        start: wordStartIndex,
                        end: wordEndIndex,
                    },
                    status: "Word (long phrase) found and wrapped without context matching",
                });
    
                // Wrap the word with the <span> tag
                return transcript.slice(0, wordStartIndex) + highlightedWord + transcript.slice(wordEndIndex);
            }
    
            // If no match is found, return the original transcript unmodified
            return transcript;
        }
    
        // Handling the special rare case: Trim context before the word and only focus on context after the word
        function handleRareCaseWithPostContext() {
            // Cut down the context up to the word
            let wordIndexInContext = cleanContext.indexOf(cleanString(word));
            if (wordIndexInContext !== -1) {
                let postWordContext = cleanContext.slice(wordIndexInContext + word.length).trim(); // Context after the word
    
                if (postWordContext.length > 0) {
                    let wordRegex = new RegExp(`\\b${word}\\b`, 'gi'); // Exact word match regex
                    let match;
    
                    while ((match = wordRegex.exec(transcript)) !== null) {
                        let wordStartIndex = match.index;
                        let wordEndIndex = wordStartIndex + word.length;
    
                        // Skip if word is already inside a span tag
                        let precedingText = transcript.slice(0, wordStartIndex);
                        let followingText = transcript.slice(wordEndIndex);
    
                        if (spanRegex.test(precedingText + word + followingText)) {
                            continue;
                        }
    
                        // Now, check if the context after the word matches with the remaining context
                        let transcriptPostWord = cleanTranscript.slice(wordEndIndex).trim();
                        if (transcriptPostWord.startsWith(postWordContext.slice(0, 10))) {  // Match a portion of the post context
                            // Log when the word is found and wrapped
                            console.log({
                                word: word,
                                context: context,
                                cleanedContext: cleanContext,
                                transcript: transcript,
                                errorId: errorId,
                                wordPosition: {
                                    start: wordStartIndex,
                                    end: wordEndIndex,
                                },
                                status: "Word (rare case with post context) found and wrapped",
                            });
    
                            // Wrap the word with the <span> tag
                            return transcript.slice(0, wordStartIndex) + highlightedWord + transcript.slice(wordEndIndex);
                        }
                    }
                }
            }
    
            return null; // Return null if no match is found
        }
    
        // Apply the rare case logic
        let rareCaseResult = handleRareCaseWithPostContext();
        if (rareCaseResult) return rareCaseResult;
    
        // Regular logic follows...
    
        // Greedily search for the word within the transcript with context consideration
        function findContextInTranscript(cleanContext, cleanTranscript) {
            let minLength = 10;  // Minimum context length allowed
            let start = 0;
            let end = cleanContext.length;
            let matchedIndex = -1;
            let bestMatch = { index: -1, similarity: 0 };
    
            // Loop through progressively smaller portions of the context to find partial matches
            while (end - start >= minLength) {
                let currentContext = cleanContext.slice(start, end).trim();
    
                if (cleanTranscript.includes(currentContext)) {
                    matchedIndex = cleanTranscript.indexOf(currentContext);
    
                    // Check the similarity percentage of the matched portion of the transcript
                    const transcriptPortion = cleanTranscript.slice(matchedIndex, matchedIndex + currentContext.length);
                    const similarity = getMatchPercentage(currentContext, transcriptPortion);
    
                    if (similarity > bestMatch.similarity) {
                        bestMatch = { index: matchedIndex, similarity };
                    }
    
                    // Break if we find a good enough match (even for partial context)
                    if (bestMatch.similarity >= matchThreshold) {
                        break;
                    }
                }
    
                // Progressively trim context from both ends for greedy search
                start += 1;
                end -= 1;
            }
    
            return bestMatch.similarity >= matchThreshold ? bestMatch.index : -1;
        }
    
        // Find the matching context in the transcript
        let contextIndex = findContextInTranscript(cleanContext, cleanTranscript);
    
        if (contextIndex === -1) {
            // console.log({
            //     word: word,
            //     context: context,
            //     transcript: transcript,
            //     errorId: errorId,
            //     status: "Word not found or insufficient context match",
            // });
            return transcript; // Return unmodified if context isn't found or match is insufficient
        }
    
        // Check if the word is within the matched context range
        let wordRegex = new RegExp(`\\b${word}\\b`, 'gi'); // Exact word match regex
        let match;
    
        while ((match = wordRegex.exec(transcript)) !== null) {
            let wordStartIndex = match.index;
            let wordEndIndex = wordStartIndex + word.length;
    
            // Skip if word is already inside a span tag
            let precedingText = transcript.slice(0, wordStartIndex);
            let followingText = transcript.slice(wordEndIndex);
    
            if (spanRegex.test(precedingText + word + followingText)) {
                continue;
            }
    
            // Check if the found word falls within the matched context
            if (wordStartIndex >= contextIndex && wordEndIndex <= contextIndex + cleanContext.length) {
                // Log when the word is found and wrapped
                console.log({
                    word: word,
                    context: context,
                    cleanedContext: cleanContext,
                    transcript: transcript,
                    errorId: errorId,
                    wordPosition: {
                        start: wordStartIndex,
                        end: wordEndIndex,
                    },
                    status: "Word found and wrapped",
                });
    
                // Wrap the word with the <span> tag
                return transcript.slice(0, wordStartIndex) + highlightedWord + transcript.slice(wordEndIndex);
            }
        }
    
        // If no match is found, return the original transcript unmodified
        return transcript;
    }                

    matches.forEach(match => {
        
        // Extract values from the match object, trimmed for safety
        const originalWord = match.context.text.substr(match.context.offset, match.length).trim();
        const replacements = match.replacements.map(replacement => `<span>${replacement.value}</span>`).join(' ');
        const detailedExplanation = match.message.trim();
        // console.log(`Context: ${match.context.text}`)
        // console.log(`Word: ${originalWord}`)
        // console.log(`Length: ${match.length}`)
        // console.log(`Offset: ${match.context.offset}`)
        // If shortMessage is empty, extract the first sentence from detailedExplanation
        let shortExplanation = match.shortMessage.trim();
        if (!shortExplanation) {
            const firstSentenceMatch = detailedExplanation.match(/[^\.!\?]+[\.!\?]+/); // Match the first sentence
            shortExplanation = firstSentenceMatch ? firstSentenceMatch[0] : '';
        }

        // Generate the HTML structure
        html += `
        <div class="grammer_error" id="GRAMMER_ERROR_${errorCounter}">
            <span class="original-grammer-word">${originalWord}</span>
            <span class="arrow">-&gt;</span> 
            <span class="improved-grammer-word">${replacements}</span>
            <span class="grammer-short-explanation"> - ${shortExplanation}</span><br>
            <span class="explanation">${detailedExplanation}</span>
        </div>`;


        // Highlight Errors In Transcript 
        // Creating the grammerErrorObj
        let grammerErrorObj = {
            originalWord: originalWord.replace(/"/g, ''), // Remove quotes if needed
            replacements: replacements,
            shortExplanation: shortExplanation.replace(/"/g, ''),
            detailedExplanation: detailedExplanation.replace(/"/g, ''),
            sentence: match.sentence,
            errorId : errorCounter,
        };

        buildGrammerErrorPopup(grammerErrorObj);

        let found = false; // Flag to track if the error word is found and wrapped

        // First Loop: Try to find and wrap the error word using context, offset, and length
        $grammertranscriptWrap.find('.file-block .transcript-text').each(function(index, element) {
            if (found) return false; // Exit loop if the word is already found
            
            let transcript = jQuery(element).html(); // This is the Transcript
        
            // Wrap error word in transcript if found using context, offset, and length
            let updatedTranscript = wrapErrorWordInTranscript(
                grammerErrorObj.originalWord,
                match.context.text,
                transcript,
                // match.context.offset,
                // match.context.length,
                grammerErrorObj.errorId
            );
        
            // Check if the error word was wrapped (i.e., the transcript was modified)
            if (updatedTranscript !== transcript) {
                // console.log(updatedTranscript);
                jQuery(element).html(updatedTranscript); // Update the HTML with the wrapped version
                console.log("Error word found using strict methods and context");
                console.table([{
                    originalWord: grammerErrorObj.originalWord,
                    contextText: match.context.text,
                    transcript: transcript,
                    offset: match.context.offset,
                    length: match.length,
                    errorId: grammerErrorObj.errorId
                }]);
                found = true; // Set flag to true to stop further iterations
            }
        });
        
        if(!found){
            console.log(`Error word "${originalWord}" Not found`);
            notFoundinTranscriptErrorIDs.push(grammerErrorObj.errorId);
        }
        errorCounter++;
    });

    // Use jQuery to insert the generated HTML into the grammar-suggestions container
    $grammerErrorsWrap.html(html);

    // Apply showAndHideElements to each grammar error div
    $grammerErrorsWrap.find('.grammer_error').each(function () {
        handleClickForGrammerError(jQuery(this));
    });
    
    
    $grammerErrorsWrap.find('.improved-grammer-word span').on('click', function(event) {
        event.stopPropagation();

        let correctWordWrap = jQuery(this);
        
        // Find the closest .grammer_error element
        let grammerErrorWrap = correctWordWrap.closest('.grammer_error');
        // Get the ID of the grammer_error element (ID format: GRAMMER_ERROR_X)
        let grammerErrorId = grammerErrorWrap.attr('id');
        
        // Convert the ID to the format ERROR_X (remove the 'GRAMMER_' part)
        let errorId = grammerErrorId.replace('GRAMMER_', '');
        
        // Find the element with ID ERROR_X
        let errorElement = jQuery(`.grammer-error#${errorId}`);
        // Remove the 'grammer-error' class from the found element
        errorElement.removeClass('grammer-error');
        
        // Get the content inside the GRAMMER_ERROR_X wrap (corrected word)
        let correctedContent = correctWordWrap.html();
        
        // Replace the content inside the .grammer-error element and make it bold
        errorElement.html(`<strong>${correctedContent}</strong>`);
        // Make the li.upgrade_grammer element disappear with a fade-out animation
        jQuery(`.grammer_error#${grammerErrorId}`).fadeOut();

        // Check if all errors have been resolved
        if ($grammertranscriptWrap.find('.grammer-error').length === 0) {
            // Append a message and hide it initially
            jQuery('#grammer-suggestions').append('<div class="grammer-resolved-message" style="display:none;">All Grammar Errors Resolved</div>');
            jQuery('.grammer-resolved-message').fadeIn(1000);
        }
    });

    addClickEventListenerToGrammerError();
    // Remove Not Found Errors from List
    notFoundinTranscriptErrorIDs.forEach(errorId => {
        document.querySelector(`#GRAMMER_ERROR_${errorId}`).remove();
    });
}


function buildGrammerErrorPopup(grammerErrorObj) {
    // Helper function to create individual popup rows
    function createPopupRow(labelText, contentText, contentClass = "orignal") {
        // Create row div
        const rowDiv = document.createElement('div');
        rowDiv.className = 'grammer-error-popup-row';

        // Create label div
        const labelDiv = document.createElement('div');
        labelDiv.className = `popup-${contentClass}-grammer-label`;
        labelDiv.textContent = labelText;

        // Create content div
        const contentDiv = document.createElement('div');
        contentDiv.className = `popup-${contentClass}-grammer`;
        contentDiv.innerHTML = `${contentText}`;

        // Append label and content to the row
        rowDiv.appendChild(labelDiv);
        rowDiv.appendChild(contentDiv);
        return rowDiv;
    }

    // Create main div for the popup
    const grammerError = document.createElement('div');
    grammerError.className = 'grammer-error-popup';
    grammerError.id = `grammer_error_popup_${grammerErrorObj.errorId}`;

    // Create inner div with dynamic id
    const innerWrap = document.createElement('div');
    innerWrap.className = 'grammer-error-inner';

    // "You Said" Row
    const originalWordRow = createPopupRow('You Said', grammerErrorObj.originalWord, 'orignal');

    // "Suggestions" Row
    const suggestionsRow = createPopupRow('Suggestions', grammerErrorObj.replacements, 'suggested');

    // "Explanation" Row
    const explanationRow = createPopupRow('Explanation', grammerErrorObj.detailedExplanation, 'exp');

    // Append rows to inner wrapper
    innerWrap.appendChild(originalWordRow);
    innerWrap.appendChild(suggestionsRow);
    innerWrap.appendChild(explanationRow);

    // Append inner wrapper to the main div
    grammerError.appendChild(innerWrap);

    // Append the main div to the body
    document.body.appendChild(grammerError);

    // Add Event Listener to the suggested replacements
    const suggestedGrammerElement = jQuery(grammerError).find('.popup-suggested-grammer span');
    suggestedGrammerElement.on('click', function (event) {
        // Extract error ID from grammerErrorObj
        const errorId = grammerErrorObj.errorId;

        // Handle the click event logic
        const originalWord = grammerErrorObj.originalWord;
        const improvedWord = this.textContent;

        // Example: Replace the original word with the improved word
        let unmarkedText = $grammertranscriptWrap.find(`.grammer-error#ERROR_${errorId}`).html();

        // Escape any special characters in the original word
        const escapedOriginalWord = originalWord.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

        // Replace the original word with the improved word and make it bold
        const updatedText = unmarkedText.replace(new RegExp(escapedOriginalWord, "gi"), `<b>${improvedWord}</b>`);
        $grammertranscriptWrap.find(`.grammer-error#ERROR_${errorId}`).html(updatedText).removeClass('grammer-error');

        // Make the li.upgrade_grammer element disappear with a fade-out animation
        jQuery(`.grammer_error#GRAMMER_ERROR_${errorId}`).fadeOut();

        // Check if all errors have been resolved
        if ($grammertranscriptWrap.find('.grammer-error').length === 0) {
            // Append a message and hide it initially
            jQuery('#grammer-suggestions').append('<div class="grammer-resolved-message" style="display:none;">All Grammar Errors Resolved</div>');
            jQuery('.grammer-resolved-message').fadeIn(1000);
        }
    });
}
/**
 * Hides and shows specific elements in the generated grammar error HTML when interacting.
 * @param {jQuery} $newDiv - The jQuery object representing the div element.
 */
function handleClickForGrammerError($newDiv) {
    const $elementsToHide = $newDiv.find(".arrow, .improved-grammer-word, .explanation");
    const $elementsToShow = $newDiv.find(".grammer-short-explanation");
    // Initially hide the elements that need to be hidden
    $elementsToHide.hide();

    $newDiv.on('click', function (event) {
        event.stopPropagation();
        let errorId = jQuery(this).attr('id');
        let transcriptErrorId = errorId.replace('GRAMMER_', '');
    
        // Get the parent container to update its max-height
        const $scrollableParent = jQuery('.suggestions-parent-wrap .e-n-carousel.swiper');
        $scrollableParent.data('MaxHeightAdded',0); // Reset the Value
        // Get Current Height of $newDiv before expansion
        const initialHeight = this.scrollHeight;
        // UnExpand Other Divs
        jQuery(".grammer_error").not(this).each(function () {
            const $thisDiv = jQuery(this);
            
            // Check if the div contains a previously added height in the data attribute
            const prevHeightDiff = $thisDiv.data('prevHeightDiff') || 0;
            // If there's a previous height difference, subtract it from the scrollable parent's max-height
            if (prevHeightDiff) {
                let currentMaxHeight = parseInt($scrollableParent.css('max-height'), 10) || 0;
                $scrollableParent.css('max-height', (currentMaxHeight - prevHeightDiff) + 'px');
                
                // Clear the stored height difference after removing it
                $thisDiv.data('prevHeightDiff', 0);
            }
    
            // Collapse the div by removing the 'expanded' class and hiding the content
            $thisDiv.removeClass('expanded');
            $thisDiv.find('.arrow, .improved-grammer-word, .explanation').slideUp(200);
            $thisDiv.find('.grammer-short-explanation').slideDown(200);
        });
    
        // Expand Current Div
        jQuery(this).addClass('expanded');
        jQuery(this).find('.grammer-short-explanation').slideUp(200);
        jQuery(this).find('.arrow, .improved-grammer-word, .explanation').slideDown(200, () => {
            // SlidDown Event Seems Triggering Function Multiple Times and We Only want to add height one time.
            if(!$scrollableParent.data('MaxHeightAdded')){
                // Get New Height of $newDiv after expansion
                const newHeight = this.scrollHeight;
                // Calculate the difference in height
                const heightDifference = newHeight - initialHeight;
        
                // Get the current max-height of the scrollable parent container
                let initMaxHeightofParent = parseInt($scrollableParent.css('max-height'), 10) || 0;

                // Update the max-height of the scrollable parent container by adding the new height difference
                const newMaxHeight = initMaxHeightofParent + heightDifference;
                $scrollableParent.data('MaxHeightAdded',1);
                $scrollableParent.css('max-height', newMaxHeight + 'px');
        
                // Store the new height difference so it can be removed later
                jQuery(this).data('prevHeightDiff', heightDifference);
            }
        });
    
        // Scroll Logic for Grammar Error
        $grammertranscriptWrap.find('.grammer-error').removeClass('grammer-highlighted');
        const $grammarErrorEl = $grammertranscriptWrap.find(`.grammer-error#${transcriptErrorId}`).addClass('grammer-highlighted');
    
        setTimeout(() => {
            if (!$grammarErrorEl.length) {
                console.error('Error: Element not found');
                return;
            }
    
            if (!$grammarErrorEl.is(':visible')) {
                console.error('Error: Element is not visible');
                return;
            }
    
            const $scrollableContainer = jQuery('.transcript-parent-wrap');
    
            if ($scrollableContainer.length) {
                const elementOffset = $grammarErrorEl.offset().top;
                const containerOffset = $scrollableContainer.offset().top;
                const scrollPosition = elementOffset - containerOffset + $scrollableContainer.scrollTop() - 50;
    
                $scrollableContainer.animate({
                    scrollTop: scrollPosition
                }, 400);
            } else {
                console.error('Error: Scrollable container not found.');
            }
        }, 200);
    });    
}

function addClickEventListenerToGrammerError() {
    $grammertranscriptWrap.find('.grammer-error').on('click', function (event) {
        event.stopPropagation(); // Stop the event from bubbling up

        const grammerErrorId = jQuery(this).attr('id');
        const grammerDivId = grammerErrorId.replace('ERROR_', 'GRAMMER_ERROR_');
        const grammerPopupId = grammerErrorId.replace('ERROR_', 'grammer_error_popup_');
        
        // Find the corresponding div with the same grammerDivId
        const $correspondingDiv = jQuery(`#${grammerDivId}`);
        const $correspondingPopup = jQuery(`#${grammerPopupId}`);
        
        jQuery('.grammer-error-popup').removeClass('active');
        jQuery('.grammer-error-popup').css({
            position: 'absolute',
            bottom: 0, // Just below the click location
            left: '-100vw',
            opacity: 0
        });
        $correspondingPopup.addClass('active');

        // Position the popup just below the click location
        const clickX = event.pageX;
        const clickY = event.pageY;
        $correspondingPopup.css({
            position: 'absolute',
            top: clickY + 10 + 'px', // Just below the click location
            left: clickX + 'px',
            opacity: 1
        }); // Optionally animate the popup appearance

        if (!$correspondingDiv.hasClass("expanded")) {
            const $scrollableParent = jQuery('.suggestions-parent-wrap .e-n-carousel.swiper');
            $scrollableParent.data('MaxHeightAdded',0); // Reset the Value
            // Get Current Height of $newDiv before expansion
            const initialHeight = $correspondingDiv[0].scrollHeight;
            // Hide elements and show the short explanation of other list items with the "upgrade_grammer" class
            // UnExpand Other Divs
            jQuery(".grammer_error").not($correspondingDiv).each(function () {
                const $thisDiv = jQuery(this);
                
                // Check if the div contains a previously added height in the data attribute
                const prevHeightDiff = $thisDiv.data('prevHeightDiff') || 0;
                
                // If there's a previous height difference, subtract it from the scrollable parent's max-height
                if (prevHeightDiff) {
                    let currentMaxHeight = parseInt($scrollableParent.css('max-height'), 10) || 0;
                    $scrollableParent.css('max-height', (currentMaxHeight - prevHeightDiff) + 'px');
                    
                    // Clear the stored height difference after removing it
                    $thisDiv.data('prevHeightDiff', 0);
                }
        
                // Collapse the div by removing the 'expanded' class and hiding the content
                $thisDiv.removeClass('expanded');
                $thisDiv.find('.grammer-short-explanation').slideDown(200);
                $thisDiv.find('.arrow, .improved-grammer-word, .explanation').slideUp(200);
            });

            // Remove existing highlights
            $grammertranscriptWrap.find(".grammer-error").removeClass("grammer-highlighted");

            // Add class to the current one
            jQuery(this).addClass("grammer-highlighted");

            // Show or hide specific elements within the clicked list item
            // Expand Current Div
            $correspondingDiv.addClass('expanded');
            $correspondingDiv.find('.grammer-short-explanation').slideUp(200);
            // $correspondingDiv.find('.arrow, .improved-grammer-word, .explanation').slideDown(200);
            $correspondingDiv.find('.arrow, .improved-grammer-word, .explanation').slideDown(200, () => {
                // SlidDown Event Seems Triggering Function Multiple Times and We Only want to add height one time.
                if(!$scrollableParent.data('MaxHeightAdded')){
                    // Get New Height of $newDiv after expansion
                    const newHeight = $correspondingDiv[0].scrollHeight;
                    // Calculate the difference in height
                    const heightDifference = newHeight - initialHeight;
            
                    // Get the current max-height of the scrollable parent container
                    let initMaxHeightofParent = parseInt($scrollableParent.css('max-height'), 10) || 0;
                    // Update the max-height of the scrollable parent container by adding the new height difference
                    const newMaxHeight = initMaxHeightofParent + heightDifference;
                    $scrollableParent.data('MaxHeightAdded',1);
                    $scrollableParent.css('max-height', newMaxHeight + 'px');
            
                    // Store the new height difference so it can be removed later
                    jQuery(this).data('prevHeightDiff', heightDifference);
                }
            });
        }
    });
}


$document.on('click', function () {
    // Get the parent container to update its max-height
    const $scrollableParent = jQuery('.suggestions-parent-wrap .e-n-carousel.swiper');
    // UnExpand Divs
    jQuery(".grammer_error").each(function () {
        const $thisDiv = jQuery(this);
        
        // Check if the div contains a previously added height in the data attribute
        const prevHeightDiff = $thisDiv.data('prevHeightDiff') || 0;
        // If there's a previous height difference, subtract it from the scrollable parent's max-height
        if (prevHeightDiff) {
            let currentMaxHeight = parseInt($scrollableParent.css('max-height'), 10) || 0;
            $scrollableParent.css('max-height', (currentMaxHeight - prevHeightDiff) + 'px');
            
            // Clear the stored height difference after removing it
            $thisDiv.data('prevHeightDiff', 0);
        }

        // Collapse the div by removing the 'expanded' class and hiding the content
        $thisDiv.removeClass('expanded');
        $thisDiv.find('.arrow, .improved-grammer-word, .explanation').slideUp(200);
        $thisDiv.find('.grammer-short-explanation').slideDown(200);
    });
    
    $grammertranscriptWrap.find(".grammer-error").removeClass("grammer-highlighted");
    jQuery('.grammer-error-popup').removeClass('active');
    jQuery('.grammer-error-popup').css({
        position: 'absolute',
        bottom:0, // Just below the click location
        left: '-100vw',
        opacity: 0
    });
});