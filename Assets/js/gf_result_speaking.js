window.onload = function() {
    jQuery(document).ready(function($) {
            // Sync Two Result Page Sliders 
            const transcriptCarousel = document.querySelector('#transcript-carousel .swiper').swiper;
            const resultDataCarousel = document.querySelector('#result-data-carousel .swiper').swiper;
            const scoresCarousel = document.querySelector('#scores-carousel .swiper').swiper;
            const audioPagination = document.getElementById('isfp-audio-pagination');

            resultDataCarousel.allowTouchMove = false;
            transcriptCarousel.allowTouchMove = false;

            transcriptCarousel.autoHeight = true;
            resultDataCarousel.autoHeight = true;

            transcriptCarousel.on('slideChange', function() {
                resultDataCarousel.slideTo(transcriptCarousel.activeIndex);
                scoresCarousel.slideTo(transcriptCarousel.activeIndex);
                console.log('Adjust Height Triggered');
                adjustDataHeight(resultDataCarousel);
            });

            resultDataCarousel.on('slideChange', function() {
                transcriptCarousel.slideTo(resultDataCarousel.activeIndex);
                scoresCarousel.slideTo(resultDataCarousel.activeIndex);
            });

            scoresCarousel.on('slideChange', function() {
                transcriptCarousel.slideTo(scoresCarousel.activeIndex);
                resultDataCarousel.slideTo(scoresCarousel.activeIndex);
            });
            disableShareBtn();


            // Logic For Streaming Response And Switching Slides
            let slideActiveIndex = 1; // 1 For Vocab, 2 For Grammer, 3 For Pronun, 4 For Fluency
            let lastEventType = 'message'; // Defaults to message Event
            let slidesCount = 4; // Number of Slides
            let userInteracted = false; // A Flag to Check if User Interacted or Not with Slides
            let loadingSavedData = true;
            let isPronunciationFeedActiveFlag = false;
            let isPronunciationProcessed = false;
            let readyToExport = false;
            let isErrorOccured = false;
            let secondLastFeedEndpoint = '';
           
            window.audioFiles = null; //Variable to Hold File Paths
            let whisperResponse = []; // From Whisper [Array of Responses]
            let vocabSuggestions = ''; // From Chat Completions
            let grammerSuggestions = null; // From Language Tool
            let pronunciationResponse = []; // From Pronunciation API
            let pronunciationFileIndex = 0;

            const vocabTranscriptWrap = $('#transcript-carousel .swiper #vocab-transcript-wrap .elementor-widget-container');
            const grammerTranscriptWrap = $('#transcript-carousel .swiper #grammer-transcript-wrap .elementor-widget-container');
            const pronunTranscriptWrap = $('#transcript-carousel .swiper #pronun-transcript-wrap .elementor-widget-container');
            const fluencyTranscriptWrap = $('#transcript-carousel .swiper #fluency-transcript-wrap .elementor-widget-container');
            const improvedAnswerTranscriptWrap = $('#transcript-carousel .swiper #improved-ans-transcript-wrap .elementor-widget-container');


            var div_index = 0, div_index_str = '';
            var vocabResponseBuffer = ""; // Buffer for holding messages For Vocab
            var improvedAnswerBuffer = "";

            var responseBuffer = '';
            var md = new Remarkable();

            const formId = gfResultSpeaking.formId;
            const entryId = gfResultSpeaking.entryId;
            const nonce = gfResultSpeaking.nonce;
            const sourceUrl = `${gfResultSpeaking.restUrl}writify/v1/event_stream_openai?form_id=${formId}&entry_id=${entryId}&_wpnonce=${nonce}`;
            const source = new EventSource(sourceUrl);
            source.addEventListener('message', handleEventStream);
            source.addEventListener('whisper', handleEventStream);
            source.addEventListener('chat/completions', handleEventStream);
            source.addEventListener('languagetool', handleEventStream);
            source.addEventListener('pronunciation', handleEventStream);
            source.addEventListener('audio_urls', handleEventStream);
            source.addEventListener('wpm', handleEventStream);
            source.addEventListener('fluency_errors', handleEventStream);
            source.addEventListener('fluency_score', handleEventStream);
            source.addEventListener('pronun_score', handleEventStream);
            source.addEventListener('vocab_score', handleEventStream);
            source.addEventListener('grammer_score', handleEventStream);
            source.addEventListener('feeds', handleEventStream);
            source.addEventListener('improved_answer', handleEventStream);

            function handleEventStream(event) {
                let eventData;
                try {
                    // Attempt to parse event.data as JSON
                    eventData = JSON.parse(event.data);
                } catch (error) {
                    // If parsing fails, assign event.data directly
                    eventData = event.data;
                }
                if (event.data == "[ALLDONE]") {
                    source.close();
                }else if(event.data == "[FIRST-TIME]") {
                    loadingSavedData = false;
                } else if (event.data.startsWith("[DIVINDEX-")) {
                    // New Feed Started

                } else if (event.data == "[DONE]") {
                    // Previous Feed Completed
                    console.log(`${lastEventType} Completed`);
                    if(lastEventType == 'chat/completions'){
                        let vocabSuggestionsWrapper = jQuery('#vocab-suggestions');
                        addUpgradeVocabClass(vocabSuggestionsWrapper);
                    }

                    // Process Pronunciation Only If All Pronunciation Data is Ready
                    if(lastEventType == 'pronunciation' && !isPronunciationProcessed){
                        if(!loadingSavedData){ // Loading First Time
                            for(i = 0; i<audioFiles.length; i++){
                                processPronunciationData(pronunciationResponse[i],i);
                            }
                        }else{
                            processSavedPronunData(pronunciationResponse);
                        }
                        isPronunciationProcessed = true;
                    }

                    // Make It Default to True again As After [DONE] New Request Will be Processed
                    loadingSavedData = true;
                    
                } else if(event.type !== 'message') {
                    // Event Contains Data to Process
                    if(event.type == 'whisper' || event.type == 'chat/completions' || event.type == 'languagetool' || event.type =='pronunciation'){
                        lastEventType = event.type;
                    }
                    // We Will Always get the Audio Files In Begining of Stream
                    if(event.type == 'audio_urls'){
                        audioFiles = JSON.parse(event.data); // Array of Audio File Paths
                        audioFiles = audioFiles.response;
                        loadAudio(audioFiles[0]);
                        renderAudioPagination(audioFiles);
                        updatePagination(audioFiles[0]);
                    }

                    else if(event.type == 'whisper'){ // Whisper Streaming Logic
                        console.log('Listening to Whisper');
                        if(eventData.hasOwnProperty('error')){
                            alert('Something Went Wrong When Loading Transcript');
                            console.log(eventData);
                            isErrorOccured = true;
                            return;
                        }
                        // Store The Transcript Data From Whisper
                        let currentWhisperResponse = JSON.parse(event.data);
                        whisperResponse.push(currentWhisperResponse);
                        let fileIndex = whisperResponse.length - 1;
                        // Extract Text with TimeStamp that should be Clickable to Play Audio
                        
                        // Also Stream The Data on Frontend Live
                        let formatedResponse = formatWhisperResponse(currentWhisperResponse,audioFiles[fileIndex]);
                        
                        // vocabTranscriptWrap.append(`<strong>File: ${fileName}</strong>:`);
                        jQuery(vocabTranscriptWrap).find('.audio-transcript-section').first().before(formatedResponse);
                        jQuery(grammerTranscriptWrap).find('.audio-transcript-section').first().before(formatedResponse);
                        jQuery(pronunTranscriptWrap).find('.audio-transcript-section').first().before(formatedResponse);
                        jQuery(fluencyTranscriptWrap).find('.audio-transcript-section').first().before(formatedResponse);                        
                        jQuery(improvedAnswerTranscriptWrap).find('.audio-transcript-section').first().before(formatedResponse);                        

                        // Add Manual Feedback Option For Pronunciation
                        setupManualFeedback(whisperResponse);
                        // Remove Skeleton Loading
                        console.log(fileIndex,audioFiles.length -1);
                        if(fileIndex == (audioFiles.length -1)){ // Last Whisper
                            jQuery(vocabTranscriptWrap).find('.audio-transcript-section').remove(); // Empty the Wrapper
                            jQuery(grammerTranscriptWrap).find('.audio-transcript-section').remove(); // Empty the Wrapper
                            jQuery(pronunTranscriptWrap).find('.audio-transcript-section').remove(); // Empty the Wrapper
                            jQuery(fluencyTranscriptWrap).find('.audio-transcript-section').remove(); // Empty the Wrapper
                            jQuery(improvedAnswerTranscriptWrap).find('.audio-transcript-section').remove(); // Empty the Wrapper
                            // Mark Filler Words
                            markFillerWords();
                        }
                        
                    }else if(event.type == 'wpm'){
                        let wpm = JSON.parse(event.data).response;
                        setWpmMeterValue(wpm);
                    }else if(event.type == 'feeds'){
                        isPronunciationFeedActiveFlag = isPronunciationFeedActive(event.data);
                        if(!isPronunciationFeedActiveFlag){ // Pronun Feed Not Active
                            console.log('Removing Pronunciation Loader');
                            removePronunLoader();
                        }
                        secondLastFeedEndpoint = eventData[eventData.length - 2].meta.endpoint;
                        console.log(secondLastFeedEndpoint);
                    }
                    else if(event.type == 'fluency_score'){
                        // Render Fluency Score
                        let scoreVal = JSON.parse(event.data).response;
                        setScore('fluency_score',scoreVal);

                    }else if(event.type == 'pronun_score'){
                        // Render Pronun Score
                        let scoreVal = JSON.parse(event.data).response;
                        setScore('pronun_score',scoreVal);

                    }else if(event.type == 'vocab_score'){
                        // Render Vocab Score
                        let scoreVal = JSON.parse(event.data).response;
                        setScore('vocab_score',scoreVal);

                    }else if(event.type == 'grammer_score'){
                        // Render Grammer Score
                        let scoreVal = JSON.parse(event.data).response;
                        setScore('grammer_score',scoreVal);
                        enableShareBtn();
                    }
                    else if(event.type == 'chat/completions'){ // Vocabulary Suggestions Streaming Logic
                        if(eventData.hasOwnProperty('error')){
                            alert('Something Went Wrong When Loading Result');
                            console.log(eventData);
                            isErrorOccured = true;
                            return;
                        }
                        // Auto Slide Logic 
                        if(!userInteracted && slideActiveIndex !== 1 && !loadingSavedData){
                            // Swipe to First Slide 
                            console.log('Slide To Vocab');
                            // resultDataCarousel.slideTo(1);
                        }
                        // There Could be chat/completions events we need to store data in One For All
                        // Parse the event data to handle both response formats
                        let response = JSON.parse(event.data);
                        console.log(response);
                        let vocabSuggestionsText;

                        // Check if response has a direct 'response' property
                        if (response.response) {
                            console.log('Response Property Found');
                            vocabSuggestionsText = response.response;
                        }
                        
                        // Check if response has 'choices' array structure
                        else if (response.choices && response.choices[0] && response.choices[0].delta && response.choices[0].delta.content) {
                            vocabSuggestionsText = response.choices[0].delta.content;
                        }

                        console.log(vocabSuggestionsText);

                        // Now `text` will contain the relevant text content in both cases
                        if (vocabSuggestionsText) {
                            
                            vocabResponseBuffer += vocabSuggestionsText; // Append the extracted text to the buffer
                            var html = md.render(vocabResponseBuffer); // Convert buffer to HTML
                            jQuery('#vocab-suggestions').find('.preloader-icon').hide();
                            let vocabSuggetsionWrap = jQuery('#vocab-suggestions');
                            vocabSuggetsionWrap.html(html); // Replace the current HTML content with the processed markdown
                        }
                    }
                    else if(event.type == 'languagetool'){ // Grammer Suggestions Streaming Logic
                        if(eventData.hasOwnProperty('error')){
                            alert('Something Went Wrong When Loading Grammer Errors');
                            console.log(eventData);
                            isErrorOccured = true;
                            jQuery('#grammer-suggestions img').replaceWith('<span>Something Went Wrong While Loading Results</span>');
                            return;
                        }
                        generateGrammerErrorsList(event.data);
                    }
                    else if(event.type == 'pronunciation'){ // pronunciation Streaming Logic
                        if(eventData.hasOwnProperty('error')){
                            alert('Something Went Wrong When Loading Pronunciation');
                            console.log(eventData);
                            isErrorOccured = true;
                            return;
                        }
                        pronunciationResponse.push(event.data);
                    }else if(event.type == 'fluency_errors'){
                        markFluencyErrors(eventData.response);
                    }else if(event.type == 'improved_answer'){
                        // Handle Error 
                        if(eventData.hasOwnProperty('error')){
                            alert('Something Went Wrong When Loading Result');
                            console.log(eventData);
                            isErrorOccured = true;
                            return;
                        }

                        // There Could be chat/completions events we need to store data in One For All
                        // Parse the event data to handle both response formats
                        let response = JSON.parse(event.data);
                        console.log(response);
                        let improvedAnswerText;

                        // Check if response has a direct 'response' property
                        if (response.response) {
                            console.log('Response Property Found');
                            improvedAnswerText = response.response;
                        }
                        
                        // Check if response has 'choices' array structure
                        else if (response.choices && response.choices[0] && response.choices[0].delta && response.choices[0].delta.content) {
                            improvedAnswerText = response.choices[0].delta.content;
                        }

                        console.log(improvedAnswerText);

                        // Now `text` will contain the relevant text content in both cases
                        if (improvedAnswerText) {
                            
                            improvedAnswerBuffer += improvedAnswerText; // Append the extracted text to the buffer
                            var html = md.render(improvedAnswerBuffer); // Convert buffer to HTML
                            jQuery('#improved-ans-wrap').find('.preloader-icon').hide();
                            let improvedAnswerWrap = jQuery('#improved-ans-wrap');
                            improvedAnswerWrap.html(html); // Replace the current HTML content with the processed markdown
                        }
                    }
                    
                }
            }

            source.onerror = function (event) {
                div_index = 0;
                source.close();
                jQuery('.error_message').css('display', 'flex');
            };

            function setScore(wrapid,scoreValue){
                scoreValue = parseFloat(scoreValue).toFixed(1);
                let wrapperId = `${wrapid}_progress_bar`;
                let wrapper = document.querySelector(`#${wrapperId}`);
                let scoreValueWrap = wrapper.querySelector(`.elementor-title #${wrapid}_wrap`);
                let progressBarWrap = wrapper.querySelector('.elementor-progress-wrapper');
                let progressEl = wrapper.querySelector('.elementor-progress-bar');
                
                let progressVal = scoreValue * 10;
                scoreValueWrap.style.opacity = 1;
                progressBarWrap.setAttribute('aria-valuenow',progressVal);
                progressEl.setAttribute('data-max',progressVal);
                progressEl.style.minWidth = `${progressVal}%`;
                console.log(scoreValue);
                if(scoreValue == 0.0){
                    scoreValueWrap.innerHTML = 'N/A';
                }else{
                    scoreValueWrap.innerHTML = scoreValue;
                }
            }

            // Utility Functions 
            function formatWhisperResponse(response, fileUrl) {
                var formattedHtml = '';
                response = response.response; // Actual Response is Inside a Response Wrapper
                if (typeof response !== 'object') {
                    response = JSON.parse(response);
                }
                
                let fileName = extractHumanReadableFileName(fileUrl);
            
                function getTranscriptEntries(segments) {
                    let entries = '';
                    let combinedText = '';
                    let startTime = segments[0].start;
                    let endTime = segments[0].end;
            
                    for (let i = 0; i < segments.length; i++) {
                        const segment = segments[i];
                        combinedText += segment.text;
                        endTime = segment.end;
            
                        // Check if this segment ends with a period, meaning it's the end of a complete sentence
                        if (/[.!?]$/.test(segment.text) || i === segments.length - 1) {
                            // Create a transcript entry for the combined segments
                            entries += `
                            <div class="transcript-entry">
                                <div class="audio-control">
                                    <div class="player-actions">
                                        <span class="play-icon" onclick="playAudioSegment()" data-start="${startTime}" data-end="${endTime}" data-audio-url="${fileUrl}">
                                            <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32" fill="none">
                                                <path d="M15.9998 29.3332C23.3636 29.3332 29.3332 23.3636 29.3332 15.9998C29.3332 8.63604 23.3636 2.6665 15.9998 2.6665C8.63604 2.6665 2.6665 8.63604 2.6665 15.9998C2.6665 23.3636 8.63604 29.3332 15.9998 29.3332Z" stroke="#FF2E2E" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                                                <path d="M13.3335 10.6665L21.3335 15.9998L13.3335 21.3332V10.6665Z" stroke="#FF2E2E" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                                            </svg>
                                        </span>
                                        <span class="pause-icon" onclick="pauseAudioSegment()" data-start="${startTime}" data-end="${endTime}" data-audio-url="${fileUrl}">
                                            <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32" fill="none">
                                                <path d="M15.9998 29.3332C23.3636 29.3332 29.3332 23.3636 29.3332 15.9998C29.3332 8.63604 23.3636 2.6665 15.9998 2.6665C8.63604 2.6665 2.6665 8.63604 2.6665 15.9998C2.6665 23.3636 8.63604 29.3332 15.9998 29.3332Z" stroke="#FF2E2E" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                                                <path d="M12 10H14.6667V22H12V10Z" stroke="#FF2E2E" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                                                <path d="M17.3333 10H20V22H17.3333V10Z" stroke="#FF2E2E" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                                            </svg>
                                        </span>
                                    </div>
                                    <div class="timestamp" data-start="${startTime}" data-end="${endTime}">${formatTime(startTime)}</div>
                                </div>
                                <p class="transcript-text">${combinedText}</p>
                            </div>
                            `;
            
                            // Reset combinedText and set startTime for the next group of segments
                            combinedText = '';
                            if (i + 1 < segments.length) {
                                startTime = segments[i + 1].start;
                            }
                        }
                    }
            
                    return entries;
                }
            
                let entriesHtml = getTranscriptEntries(response.segments);
                formattedHtml = `
                    <div class="file-block">
                        <h3 class="file-title">${fileName}</h3>
                        ${entriesHtml}
                    </div>
                `;
            
                // Return the formatted HTML
                return formattedHtml;
            }

            // Helper function to format time in MM:SS format
            function formatTime(seconds) {
                var hours = Math.floor(seconds / 3600);
                var minutes = Math.floor((seconds % 3600) / 60);
                var remainingSeconds = Math.floor(seconds % 60);
                
                return (hours < 10 ? '0' : '') + hours + ':' +
                       (minutes < 10 ? '0' : '') + minutes + ':' +
                       (remainingSeconds < 10 ? '0' : '') + remainingSeconds;
            }

            function extractHumanReadableFileName(url) {
                // Step 1: Extract the file name from the URL
                var fileName = url.split('/').pop().split('#')[0].split('?')[0];
            
                // Step 2: Remove the file extension (optional)
                var nameWithoutExtension = fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
            
                // Step 3: Convert to human-readable format
                var humanReadableName = nameWithoutExtension
                    .replace(/[-_]/g, ' ')          // Replace hyphens and underscores with spaces
                    .replace(/\b\w/g, function(l) { // Capitalize the first letter of each word
                        return l.toUpperCase();
                    });
            
                return humanReadableName;
            }

            async function playAudioSegment(){
                let trigger = this.event.target;
                let start = trigger.dataset.start;
                let end = trigger.dataset.end;
                let audio = trigger.dataset.audioUrl;
                if(audio != window.loadedAudio){
                    await loadAudio(audio);
                }
                playSegment(start, end);
            }

            window.playAudioSegment = playAudioSegment;

            function renderAudioPagination(audioFiles){
                if(audioFiles.length < 2){
                    return;
                }

                const paginationEl = document.getElementById('isfp-audio-pagination');

                let paginationWrap = document.createElement('div');
                paginationWrap.classList.add('isfp-audio-pagination-wrap');
                audioFiles.forEach((audioFile,index) => {
                    let fileNumber = index + 1;
                    let pageNumberEl = document.createElement('span');
                    pageNumberEl.classList.add('audio-file-page-number');
                    pageNumberEl.textContent = fileNumber;
                    pageNumberEl.dataset.audioUrl = audioFile;
                    pageNumberEl.addEventListener('click', async ()=>{
                        let trigger = this.event.target;
                        let audioFile = trigger.dataset.audioUrl;
                        await loadAudio(audioFile);
                    })
                    paginationWrap.appendChild(pageNumberEl);
                });

                paginationEl.appendChild(paginationWrap);
            }

            function updatePagination(audioUrl) {
                const $paginationEl = $('#isfp-audio-pagination');
                
                // Remove 'active' class from all spans
                $paginationEl.find('span').removeClass('active');
                
                // Find the span with the matching data-audio-url and add 'active' class
                const $audioSpan = $paginationEl.find(`span[data-audio-url="${audioUrl}"]`);
                
                if ($audioSpan.length) {
                    $audioSpan.addClass('active');
                }
            }

            window.updatePagination = updatePagination;


            function generateGrammerErrorsList(langugetoolResponse){
                let errorResponse = JSON.parse(langugetoolResponse);
                if (errorResponse.response && errorResponse.response.errors) {
                    let errors = errorResponse.response.errors;
                    alert('Something Went Wrong While Checking For Grammer Errors');
                    console.log(errors);
                    return;
                }
                let response = JSON.parse(JSON.parse(langugetoolResponse).response);
                
                let matches = response.matches;
                // Helper Function locacted in gf_result_grammer_interaction 
                generateGrammarErrorHTML(matches);
            }

            function setWpmMeterValue(wpm) {
                const wpmArc = document.getElementById('wpmArc');
                const wpmValue = document.getElementById('wpmValue');
                const wpmStatus = document.getElementById('wpmStatus');
            
                const circumference = 126; // Approximate length of the half-circle
            
                // Ensure the WPM value is within the allowed range (0 to 300)
                const clampedWpm = Math.min(Math.max(wpm, 0), 300);
            
                // Map the WPM value (0 to 300) to a percentage (0 to 100)
                const progress = (clampedWpm / 300) * 100;
            
                // Calculate the stroke offset based on the progress (0 - 100)
                const offset = circumference - (progress / 100 * circumference);
            
                // Update the stroke-dashoffset to control the length of the arc
                wpmArc.style.strokeDashoffset = offset;
            
                // Change the stroke color based on the WPM value
                if (clampedWpm < 120 ) {
                    wpmArc.style.stroke = 'rgba(255, 73, 73, 1)'; // Red if less than 110 WPM
                    // Update the WPM value in the center of the half-circle
                    wpmValue.innerHTML = `<span style="color:rgba(255, 73, 73, 1)">${clampedWpm} WPM <br> Too Slow</span>`;
                    wpmStatus.innerHTML = `Your Pace was a little <span style="color:rgba(255, 73, 73, 1)"> Too Slow</span> during this recording`;
                } else if (clampedWpm > 180) {
                    wpmArc.style.stroke = 'rgba(255, 130, 46, 1)'; // Orange if greater than 150 WPM
                    // Update the WPM value in the center of the half-circle
                    wpmValue.innerHTML = `<span style="color:rgba(255, 130, 46, 1)">${clampedWpm} WPM <br> Too Fast</span>`;
                    wpmStatus.innerHTML = `Your Pace was a little <span style="color:rgba(255, 130, 46, 1)"> Too Fast</span> during this recording`;
                } else {
                    wpmArc.style.stroke = 'rgba(19, 206, 102, 1)'; // Green between 110 and 160 WPM
                    // Update the WPM value in the center of the half-circle
                    wpmValue.innerHTML = `<span style="color:rgba(19, 206, 102, 1)">${clampedWpm} WPM <br> Good Job</span>`;
                    wpmStatus.innerHTML = `Your Pace was <span style="color:rgba(19, 206, 102, 1)"> Good </span> during this recording`;
                }
            }


            function calculateWPM(whisperResponse) {
                let totalWpm = 0;
                let responseCount = 0;
                let wpm = 0;
                whisperResponse.forEach(responseObj => {
                    let segments = responseObj.response.segments;
                    let responseWpmTotal = 0;
                    let segmentCount = 0;
            
                    segments.forEach(segment => {
                        let start = segment.start;
                        let end = segment.end;
                        let text = segment.text.trim();
            
                        let duration = end - start; // Duration in seconds
                        let words = text.split(/\s+/).length; // Count words
                        let segmentWpm = (words / (duration / 60)); // Calculate WPM for the segment
            
                        responseWpmTotal += segmentWpm;
                        segmentCount++;
                    });
            
                    if (segmentCount > 0) {
                        totalWpm += (responseWpmTotal / segmentCount); // Average WPM per response
                        responseCount++;
                    }
                });
            
                wpm =  responseCount > 0 ? totalWpm / responseCount : 0; // Average WPM across all responses
                wpm = Math.floor(wpm);
                return wpm;
            }


            function isPronunciationFeedActive(feedDataString) {
                try {
                    // Parse the input string as JSON
                    const feedData = JSON.parse(feedDataString);
                    
                    // Loop through the feeds to find the one with the endpoint 'pronunciation'
                    for (let feed of feedData) {
                        if (feed.meta.endpoint === "pronunciation") {
                            // Check if the feed is active (is_active: "1" means active)
                            return feed.is_active === "1";
                        }
                    }
            
                    // If no feed with the endpoint 'pronunciation' is found, return false
                    return false;
                } catch (error) {
                    console.error("Error parsing feed data:", error);
                    return false;
                }
            }

            function markFluencyErrors(errors){
                let fluencyFileBlocks = $fluencytranscriptWrap.find('.file-block'); // Find all file blocks

                // Loop through each error in the errors array
                if(!errors){
                    return;
                }
                errors.forEach(error => {
                    let found = false; // Track if the word pair has been found

                    // Loop through each file block
                    fluencyFileBlocks.each(function () {
                        if (found) return; // Exit the loop if the pair is already found

                        let currentFluencyFileBlock = $(this); // Get the current file block
                        let fluencytranscripts = currentFluencyFileBlock.find('.transcript-text');

                        // Loop through each transcript text in the file block
                        fluencytranscripts.each(function () {
                            if (found) return; // Exit the loop if the pair is already found

                            let transcript = jQuery(this);
                            let transcriptText = transcript.html(); // Get the transcript text (HTML)

                            // Build the regex pattern to find the previousWord followed by the currentPronunWord
                            let regexPattern = new RegExp(`\\b${error.previousWord}\\s*[\\.,!?]?\\s*${error.currentPronunWord}\\b`, 'i');

                            // Check if the word pair is found
                            if (regexPattern.test(transcriptText)) {
                                let svgIcon = '';

                                // Determine the type of pause error and generate the appropriate SVG
                                if (error.pauseError === 'UnexpectedBreak') {
                                    svgIcon = `<svg width="39" height="28" viewBox="0 0 39 28" fill="none" xmlns="http://www.w3.org/2000/svg">
                                                    <path fill-rule="evenodd" clip-rule="evenodd" d="M5.85352 11.0264V0H0.609375V27.2969H5.85351L5.85352 16.2705H33.1504V27.2969H38.3945V0H33.1504V11.0264H5.85352Z" fill="#FF4949"></path>
                                                </svg>`;
                                } else if (error.pauseError === 'MissingBreak') {
                                    svgIcon = `<svg width="39" height="28" viewBox="0 0 39 28" fill="none" xmlns="http://www.w3.org/2000/svg">
                                                    <path fill-rule="evenodd" clip-rule="evenodd" d="M6.13867 11.0264V0H0.894531V27.2969H6.13867L6.13867 16.2705H33.4355V27.2969H38.6797V0H33.4355V11.0264H6.13867Z" fill="#FFCC3D"></path>`;
                                }

                                // Update the transcript with the pause error and SVG
                                let updatedTranscriptText = transcriptText.replace(regexPattern, `${error.previousWord} <span class="pause-error ${error.pauseError.toLowerCase()}-pause">
                                                                ${svgIcon}
                                                            </span> ${error.currentPronunWord}`);

                                // Update the transcript with the new HTML
                                transcript.html(updatedTranscriptText);

                                found = true; // Mark the error as found
                                return false; // Break out of the current transcript loop
                            }
                        });
                    });
                });
            }

            function disableShareBtn(){
                let shareBtn = document.querySelector('#share-result-trigger');
                shareBtn.querySelector('.elementor-button').style.backgroundColor = '#eee';
                shareBtn.querySelector('.elementor-button').style.pointerEvents = 'none';
            }
            function enableShareBtn(){
                let shareBtn = document.querySelector('#share-result-trigger');
                shareBtn.querySelector('.elementor-button').style.backgroundColor = '';
                shareBtn.querySelector('.elementor-button').style.pointerEvents = '';
            }

            function adjustDataHeight(resultDataCarousel) {
                let maxHeight = window.innerHeight - 300;
                let suggestionsParentWrap = document.querySelector('.suggestions-parent-wrap');
                let activeMistakeContainer = resultDataCarousel.wrapperEl.querySelector('.swiper-slide-active .mistake-container');
            
                // Log current window and maximum height values
                console.log(`Current Window Height: ${window.innerHeight}`);
                console.log(`Max Height for Mistake Container: ${maxHeight}px`);
                console.log(`Active Mistake Container Scroll Height: ${activeMistakeContainer.scrollHeight}px`);
            
                // Log current overflow state before applying changes
                console.log(`Current Overflow State of suggestionsParentWrap: ${window.getComputedStyle(suggestionsParentWrap).overflowY}`);
            
                // Check if the content height exceeds the max allowed height
                if (activeMistakeContainer.scrollHeight > maxHeight) {
                    console.log('Applying Overflow: auto (Container height exceeds maxHeight)');
                    suggestionsParentWrap.style.setProperty('overflow-y', 'auto', 'important');
                } else {
                    console.log('Applying Overflow: hidden (Container height is within limits)');
                    suggestionsParentWrap.style.setProperty('overflow-y', 'hidden', 'important');
                }
            
                // Get the previous height of the Swiper container
                let previousHeight = suggestionsParentWrap.querySelector('.e-n-carousel.swiper').offsetHeight;
            
                // Calculate the new height
                let newHeight = activeMistakeContainer.scrollHeight + 20;
            
                // Log if the height increased or decreased
                if (newHeight > previousHeight) {
                    console.log(`Height of e-n-carousel increased from ${previousHeight}px to ${newHeight}px`);
                } else if (newHeight < previousHeight) {
                    console.log(`Height of e-n-carousel decreased from ${previousHeight}px to ${newHeight}px`);
                } else {
                    console.log(`Height of e-n-carousel remained the same at ${newHeight}px`);
                }
            
                // Apply the new max height to the Swiper container
                suggestionsParentWrap.querySelector('.e-n-carousel.swiper').style.setProperty('max-height', `${newHeight}px`, 'important');
            
                // Log the final state of the Swiper container after the height change
                let finalHeight = suggestionsParentWrap.querySelector('.e-n-carousel.swiper').offsetHeight;
                console.log(`Final Swiper Container Height After Change: ${finalHeight}px`);
            
                // Log the element to inspect it if needed
                console.log(suggestionsParentWrap.querySelector('.e-n-carousel.swiper'));
            }                        
    });
};