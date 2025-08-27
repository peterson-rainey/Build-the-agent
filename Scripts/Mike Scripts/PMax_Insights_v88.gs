// v88 - minor improvements over v86
// template sheet to copy -> https://docs.google.com/spreadsheets/d/1pbF1ndCmGP5gl6D9Ynp7SN1bTVI0cckLmC3XUzYpNao/copy 


const SHEET_URL = ''         // create a copy of the template above first (otherwise a new sheet is created every time)

const CLIENTCODE = ''        // this string will be added to your sheet name to make sheet management easier. It doesn't affect the running of the script at all.



// wiki with instructions -> https://pmax.super.site/


// if you want to see the logs, set this to true
const DEBUG_LOGS_ON = false;

// please don't change any code below this line, thanks! ——————————————————————————————————————————————————————————————————————————————



const OTHER_SETTINGS = {
    scriptVersion: 'v88',
    minLpImpr: 50,
    weeklyDays: 366,
    errorCol: 14,
    comingFrom: 'single',
    mccTimezone: AdsApp.currentAccount().getTimeZone(),
    defaultSettings: { numberOfDays: 30, tCost: 10, tRoas: 4, minPlaceImpr: 5, brandTerm: '', accountType: false },
    urls: {
        clientNew: typeof SHEET_URL !== 'undefined' ? SHEET_URL : '',
        template: 'https://docs.google.com/spreadsheets/d/1pbF1ndCmGP5gl6D9Ynp7SN1bTVI0cckLmC3XUzYpNao/', // please don't change these URLs
        backup: 'https://docs.google.com/spreadsheets/d/1pbF1ndCmGP5gl6D9Ynp7SN1bTVI0cckLmC3XUzYpNao/',
        dgTemplate: 'https://docs.google.com/spreadsheets/d/1ohhff-XM5eBP0Umto4Fg_kifRxluLqQ-ADzvVkaIn18/',
        aiTemplate: 'https://docs.google.com/spreadsheets/d/1Lduw-ZhKcNUPN4CN2iqU3v_rU-PeXq_wXNy2MnJulM4/',
    },
};
const RUN_ASSET_FATIGUE_ANALYSIS = false;

function main() {

    let clientSettings = { whispererUrl: '', dgUrl: '', brandTerm: '', accountType: 'ecommerce' };
    let sheetUrl = typeof SHEET_URL !== 'undefined' ? SHEET_URL : '';
    let clientCode = CLIENTCODE;
    let mcc = OTHER_SETTINGS;

    Logger.log('Starting the PMax Insights script.');
    const mainTasksStart = new Date();

    let { ss, s, start, ident } = configureScript(sheetUrl, clientCode, clientSettings, mcc);

    s.timezone = AdsApp.currentAccount().getTimeZone();
    let { mainDateRange, dateRangeString, fromDate, toDate } = prepareDateRange(s);
    let elements = defineElements(s);
    let queries = buildQueries(elements, s, dateRangeString);

    if (DEBUG_LOGS_ON) Logger.log('Fetching campaign data. This may take a few moments.');
    const dataFetchStart = new Date();
    let data = getData(queries, s);
    if (DEBUG_LOGS_ON) Logger.log(`Campaign data fetched in ${((new Date() - dataFetchStart) / 1000).toFixed(2)} seconds`);

    let dgResults = null;

    if (data !== null) {
        if (DEBUG_LOGS_ON) Logger.log('Starting data processing.');
        const dataProcessingStart = new Date();
        let { dAssets, dAGData, dTotal, dSummary, terms, totalTerms, sNgrams, tNgrams, placements, idCount, dShoppingProducts } = processAndAggData(data, ss, s, mainDateRange);
        if (DEBUG_LOGS_ON) Logger.log(`Data processing completed in ${((new Date() - dataProcessingStart) / 1000).toFixed(2)} seconds`);

        if (DEBUG_LOGS_ON) Logger.log('Writing data to sheets.');
        const sheetWriteStart = new Date();
        writeAllDataToSheets(ss, { dAssets, dAGData, dTotal, dSummary, terms, totalTerms, sNgrams, tNgrams, placements, idCount, dShoppingProducts, ...data });
        processAdvancedSettings(ss, s, queries, dateRangeString, elements);
        if (DEBUG_LOGS_ON) Logger.log(`Sheet writing completed in ${((new Date() - sheetWriteStart) / 1000).toFixed(2)} seconds`);

        clientSettings.totalData = dTotal;
    } else {
        clientSettings.message = 'No eligible PMax/Shopping campaigns found.';
    }

    if (s.turnonDemandGen) {
        if (DEBUG_LOGS_ON) Logger.log('Starting Demand Gen processing.');
        const dgStart = new Date();

        const dgCheckQuery = 'SELECT campaign.id FROM campaign WHERE ' + dateRangeString + elements.dgOnly + elements.campLike + ' LIMIT 1';
        if (fetchData(dgCheckQuery).length > 0) {
            dgResults = processDG(ss, s, ident, mcc, dateRangeString, elements, fromDate, toDate);
            clientSettings.dgData = dgResults.dgSummary;
            clientSettings.dgUrl = dgResults.dgUrl;
            if (DEBUG_LOGS_ON) Logger.log(`Demand Gen processing completed in ${((new Date() - dgStart) / 1000).toFixed(2)} seconds`);
        } else {
            clientSettings.message = 'No eligible Demand Gen campaigns found.';
            if (DEBUG_LOGS_ON) Logger.log('No eligible Demand Gen campaigns found.');
        }
    }

    if (!data && !dgResults) {
        clientSettings.message = 'No eligible PMax/Shopping/DG campaigns found.';
    }


    if (DEBUG_LOGS_ON) Logger.log('Processing AI sheet.');
    const aiStart = new Date();
    let { whispererUrl, lastRunAIDate } = processAISheet(ss, s, ident, s.aiRunAt, mcc, start);
    if (DEBUG_LOGS_ON) Logger.log(`AI processing completed in ${((new Date() - aiStart) / 1000).toFixed(2)} seconds`);

    adjustSheetVisibilityBasedOnAccountType(ss, s);
    log(ss, start, s, ident);


    /*
    // MCC snippet 
    if (!clientSettings.totalData && !clientSettings.dgData) {
        Logger.log('No PMax/Shopping or DG data to return to MCC script.');
        // still return urls and messages
    }

    clientSettings.clientUrl = ss.getUrl() || '';
    clientSettings.whispererUrl = whispererUrl || '';
    clientSettings.dgUrl = clientSettings.dgUrl || '';
    clientSettings.lastRunAIDate = lastRunAIDate || 'no last run date';
    ss.getSheetByName('MCC Script').hideSheet();
    ss.getSheetByName('Save $100').hideSheet();

    const mainTasksTime = ((new Date() - mainTasksStart) / 1000).toFixed(2);
    if (DEBUG_LOGS_ON) Logger.log(`Total main tasks execution time: ${mainTasksTime} seconds`);

    return clientSettings;
    */

} // end func

//#region AI & DG functions -------------------------------------------------------------------------
function processAISheet(ss, s, clientCode, aiRunAt, mcc, start) {

    let whispererUrl = s.whispererUrl;
    let aiTemplate = s.aiTemplate;

    // start whisperer main

    let lastRunAIDate = '';

    try {
        let start = new Date();
        if (DEBUG_LOGS_ON) Logger.log('Processing AI sheet.');
        const aiProcessStart = new Date();

        if (!whispererUrl) {
            Logger.log('No AI sheet URL found. Creating one from template.')
            let aiSheet = safeOpenAndShareSpreadsheet(aiTemplate, true, `${clientCode} - AI Whisperer ${mcc.scriptVersion} - MikeRhodes.com.au (c)`);
            whispererUrl = aiSheet.getUrl();
            Logger.log(`New ${mcc.scriptVersion} of AI Whisperer sheet created: ${whispererUrl}`);
        }

        let aiSheet = safeOpenAndShareSpreadsheet(whispererUrl);
        updateCrossReferences(ss, aiSheet, whispererUrl);

        if (aiRunAt === 99) {
            Logger.log(`AI Run At Day is set to 'Never', so skipping AI section. Alter settings in the sheet to change this.`);
            lastRunAIDate = 'AI Run At Day is set to Never';
            return { whispererUrl, lastRunAIDate };
        } else if (mcc.mccDay !== aiRunAt && aiRunAt !== -1) {
            Logger.log(`AI Run At Day is not today. Not Processing AI section.`);
            lastRunAIDate = 'AI Run ≠ today';
            return { whispererUrl, lastRunAIDate };
        }

        let aiSet = getSettingsFromAISheet(aiSheet);
        aiSet = enrichSettings(aiSet, whispererUrl, s, ss, aiSheet);

        let reportsToGenerate = getReportsToGenerate(aiSet);

        if (!reportsToGenerate.length) {
            Logger.log('No reports to generate. Exiting AI section.');
            lastRunAIDate = 'No reports to generate.'; // return useful message for MCC sheet
            return { whispererUrl, lastRunAIDate };
        }

        if (!aiSet.apiKey && !aiSet.anth_apikey) {
            Logger.log('No API keys found. Exiting AI section.');
            lastRunAIDate = 'No API keys found.'; // return useful message for MCC sheet
            return { whispererUrl, lastRunAIDate };
        }

        let { endpoint, headers, model } = initializeModel(aiSet);
        aiSet.modelOut = model;

        processReports(reportsToGenerate, aiSet, endpoint, headers, model, aiSheet, mcc);

        Logger.log('Finished AI Process. Total duration: ' + ((new Date() - start) / 1000).toFixed(0) + ' seconds');
        Logger.log('Total cost: ' + aiSet.totalCost.toFixed(2) + '\n');

        lastRunAIDate = Utilities.formatDate(new Date(), mcc.mccTimezone, "MMMM-dd HH:mm");
    } catch (error) {
        Logger.log('An error occurred: ' + error.toString());
        lastRunAIDate = 'Error occurred: ' + error.toString();
    }

    const aiProcessTime = ((new Date() - aiProcessStart) / 1000).toFixed(2);
    if (DEBUG_LOGS_ON) Logger.log(`AI processing completed in ${aiProcessTime} seconds`);

    return { whispererUrl, lastRunAIDate };

    // helper funcs for AI stuff - inside processAISheet/main

    function logAndReturn(aiRunAt, whispererUrl, lastRunAIDate) {
        if (aiRunAt === 99) {
            Logger.log(`AI Run At Day is set to 'Never', so skipping AI section. Check settings on MCC sheet to change this.`);
        } else {
            Logger.log(`AI Run At Day is not today. Not Processing AI section.`);
        }
        return { whispererUrl, lastRunAIDate };
    }

    function updateCrossReferences(mainSheet, aiSheet, whispererUrl) {
        try {
            mainSheet.getRangeByName('whispererUrl').setValue(whispererUrl);
        } catch (e) {
            Logger.log(`Error setting whisperUrl in MCC sheet: ${e.message}`);
        }
        try {
            aiSheet.getRangeByName('pmaxUrl').setValue(mainSheet.getUrl());
        } catch (e) {
            Logger.log(`Error setting pmaxUrl in AI sheet: ${e.message}`);
        }
    }

    function getAISheet(whispererUrl, aiTemplate, clientCode, scriptVersion) {
        let aiSheet;
        if (whispererUrl) {
            aiSheet = safeOpenAndShareSpreadsheet(whispererUrl);
        } else {
            aiSheet = safeOpenAndShareSpreadsheet(aiTemplate, true, `${clientCode} - AI Whisperer ${scriptVersion} - MikeRhodes.com.au (c)`);
            Logger.log(`New ${scriptVersion} of AI Whisperer sheet created: ${aiSheet.getUrl()}`);
        }
        return aiSheet;
    }

    function enrichSettings(aiSet, whispererUrl, s, ss, aiSheet) {
        let enrichedSettings = {
            ...aiSet,
            whispererUrl,
            ident: s.ident,
            timezone: s.timezone,
            totalCost: 0
        };

        if (ss) {
            aiSheet.getRangeByName('pmaxUrl').setValue(ss.getUrl());
            ss.getRangeByName('whispererUrl').setValue(whispererUrl);
        } else {
            aiSheet.getSheetByName('Get the PMax Script').showSheet();

            if (!enrichedSettings.pmaxUrl) {
                Logger.log('A Pmax Insights URL should be added to the Whisperer Sheet for best results.');
                Logger.log('Note: Without a valid Pmax Insights URL, you\'ll only be able to use the \'myData\' option.');
                enrichedSettings.myDataOnly = true;
            }
        }
        return enrichedSettings;
    }

    function processReports(reportsToGenerate, aiSet, endpoint, headers, model, aiSheet, mcc) {
        for (let report of reportsToGenerate) {
            let startLoop = new Date();
            let data = getDataForReport(report, aiSet.pmaxUrl, aiSet, mcc);
            if (data === 'No data available for this report.') {
                Logger.log('No data available for ' + report + '.\nCheck the settings in the PMax Insights Sheet.');
                continue;
            }
            let { prompt, suffix, usage } = getPrompt(report, data, aiSet);
            let initialPrompt = suffix + prompt + '\n' + usage + '\n' + data;
            let { response, cost } = getReponseAndCost(initialPrompt, endpoint, headers, model, 'report generation');
            aiSet.cost = cost;
            let audio = aiSet.useVoice ? textToSpeechOpenAI(response, report, aiSet) : null;

            // if using expert mode, check if using expert prompt & rewrite whichever prompt used after eval, then run expert prompt on data
            if (aiSet.expertMode) {
                Logger.log('-- Using expert mode for ' + report);
                let { rewriteResponse, expertResponse, expertCost } = runExpertMode(aiSet, report, data, prompt, suffix, usage, response, endpoint, headers, model);
                aiSet.expertCost = expertCost;
                outputToSheet(aiSheet, expertResponse, report, audio, aiSet, model, rewriteResponse, expertCost);
            } else {
                outputToSheet(aiSheet, response, report, audio, aiSet, model);
            }

            if (aiSet.useEmail) {
                sendEmail(aiSet.email, report, response, aiSet.useVoice, audio);
            }

            let endLoop = new Date();
            let aiReportDuration = (endLoop - startLoop) / 1000;
            logAI(aiSheet, report, aiReportDuration, aiSet);
        } // loop through reports
    }

    function runExpertMode(aiSet, report, data, prompt, suffix, usage, response, endpoint, headers, model) {
        model = aiSet.llm === 'openai' ? 'o4-mini-2025-04-16' : 'claude-sonnet-4-20250514'; // use best available model with valid API key

        // 3 stages: evaluate output, rewrite, run expert prompt
        let evalPrompt = aiSet.p_evalOutput + '\n\n Original prompt:\n' + suffix + prompt + '\n' + usage + '\n\n Original data:\n' + data + '\n\n Original response:\n' + response;
        let { response: evalResponse, cost: evalCost } = getReponseAndCost(evalPrompt, endpoint, headers, model, 'evaluation');

        // rewrite original prompt
        let rewritePrompt = aiSet.p_expertMode + '\n\n Original prompt:\n' + prompt + '\n\n' + evalResponse;
        let { response: rewriteResponse, cost: rewriteCost } = getReponseAndCost(rewritePrompt, endpoint, headers, model, 'prompt rewrite');

        // run using expert prompt to get better output
        let expertPrompt = rewriteResponse + '\n\n Original prompt:\n' + suffix + prompt + '\n' + usage + '\n\n' + data;
        let { response: expertResponse, cost: eCost } = getReponseAndCost(expertPrompt, endpoint, headers, model, 'creation of better report');

        let expertCost = (parseFloat(evalCost) + parseFloat(rewriteCost) + parseFloat(eCost));
        return { rewriteResponse, expertResponse, expertCost };
    }

    function getSettingsFromAISheet(aiSheet) {
        Logger.log('Getting settings from AI sheet');
        let settingsKeys = [
            // main settings
            'llm', 'model', 'apiKey', 'anth_apikey', 'pmaxUrl', 'lang', 'expertMode', 'whoFor',
            'useVoice', 'voice', 'folder', 'useEmail', 'email', 'maxResults',
            // prompts
            'p_productTitles', 'p_landingPages', 'p_changeHistory', 'p_searchCategories',
            'p_productMatrix', 'p_nGrams', 'p_nGramsSearch', 'p_asset', 'p_placement', 'p_myData',
            'p_internal', 'p_client', 'p_expertMode', 'p_evalOutput',
            // responses
            'r_productTitles', 'r_landingPages', 'r_changeHistory', 'r_searchCategories',
            'r_productMatrix', 'r_nGrams', 'r_nGramsSearch', 'r_asset', 'r_placement', 'r_myData',
            // expert mode prompts
            'e_productTitles', 'e_landingPages', 'e_changeHistory', 'e_searchCategories',
            'e_productMatrix', 'e_nGrams', 'e_nGramsSearch', 'e_asset',
            // expert mode 'use'
            'use_productTitles', 'use_landingPages', 'use_changeHistory', 'use_searchCategories',
            'use_productMatrix', 'use_nGrams', 'use_nGramsSearch', 'use_asset'
        ];

        let settings = {};
        settingsKeys.forEach(key => {
            settings[key] = aiSheet.getRange(key).getValue();
        });

        // if no apiKey or anth_apikey found in aiSheet, then use mcc keys & write values to sheet
        if (!settings.apiKey && mcc.apikey) {
            settings.apiKey = mcc.apikey;
            aiSheet.getRangeByName('apiKey').setValue(mcc.apikey);
            Logger.log(`No API Key found in AI Sheet, using MCC API Key`);
        }
        if (!settings.anth_apikey && mcc.anth_apikey) {
            settings.anth_apikey = mcc.anth_apikey;
            aiSheet.getRangeByName('anth_apikey').setValue(mcc.anth_apikey);
            Logger.log(`No Anthropic API Key found in AI Sheet, using MCC API Key`);
        }

        return settings;
    }

    function initializeModel(aiSet) {
        Logger.log('Initializing language model.');
        let endpoint, headers, model;
        if (aiSet.llm === 'openai') {
            if (!aiSet.apiKey) {
                console.error('Please enter your OpenAI API key in the Settings tab.');
                throw new Error('Error: OpenAI API key not found.');
            }
            endpoint = 'https://api.openai.com/v1/chat/completions';
            headers = { "Authorization": `Bearer ${aiSet.apiKey}`, "Content-Type": "application/json" };
            model = aiSet.model === 'better' ? 'o4-mini-2025-04-16' : 'gpt-4.1-nano-2025-04-14';
        } else if (aiSet.llm === 'anthropic') {
            if (!aiSet.anth_apikey) {
                console.error('Please enter your Anthropic API key in the Settings tab.');
                throw new Error('Error: Anthropic API key not found.');
            }
            endpoint = 'https://api.anthropic.com/v1/messages';
            headers = { "x-api-key": aiSet.anth_apikey, "Content-Type": "application/json", "anthropic-version": "2023-06-01" };
            model = aiSet.model === 'better' ? 'claude-sonnet-4-20250514' : 'claude-3-5-haiku-20241022';
        } else {
            console.error('Invalid model indicator. Please choose between "openai" and "anthropic".');
            throw new Error('Error: Invalid model indicator provided.');
        }
        return { endpoint, headers, model };
    }

    function getReponseAndCost(prompt, endpoint, headers, model, stage) {
        if (prompt === "No data available for this report.") {
            return { response: prompt, tokenUsage: 0 };
        }

        // Setup payload based on the model and prompt
        let payload = {
            model: model,
            messages: [{ "role": "user", "content": prompt }],
            ...(model.includes('claude') && { "max_tokens": 1000 })  // Example of conditionally adding properties
        };

        let { response, inputTokens, outputTokens } = genericAPICall(endpoint, headers, payload, stage);
        let cost = calculateCost(inputTokens, outputTokens, model);

        return { response, cost };
    }

    function genericAPICall(endpoint, headers, payload, stage) {
        Logger.log(`Making API call for ${stage}.`);
        let httpOptions = {
            "method": "POST",
            "muteHttpExceptions": true,
            "headers": headers,
            "payload": JSON.stringify(payload)
        };

        let attempts = 0;
        let response;
        do {
            response = UrlFetchApp.fetch(endpoint, httpOptions);
            if (response.getResponseCode() === 200) {
                break;
            }
            Utilities.sleep(2000 * attempts); // Exponential backoff
            attempts++;
        } while (attempts < 3);

        let responseCode = response.getResponseCode();
        let responseContent = response.getContentText();

        if (responseCode !== 200) {
            Logger.log(`API request failed with status ${responseCode}. Use my free scripts to test your API key: https://github.com/mikerhodesideas/free`);
            try {
                let errorResponse = JSON.parse(responseContent);
                Logger.log(`Error details: ${JSON.stringify(errorResponse.error.message)}`);
                return { response: `Error: ${errorResponse.error.message}`, inputTokens: 0, outputTokens: 0 };
            } catch (e) {
                Logger.log('Error parsing API error response.');
                return { response: 'Error: Failed to parse the API error response.', inputTokens: 0, outputTokens: 0 };
            }
        }

        let responseJson = JSON.parse(response.getContentText());
        let inputTokens;
        let outputTokens;

        if (endpoint.includes('openai.com')) {
            return { response: responseJson.choices[0].message.content, inputTokens: responseJson.usage.prompt_tokens, outputTokens: responseJson.usage.completion_tokens };
        } else if (endpoint.includes('anthropic.com')) {
            return { response: responseJson.content[0].text, inputTokens: responseJson.usage.input_tokens, outputTokens: responseJson.usage.output_tokens };
        }
    }

    function outputToSheet(aiSheet, output, report, audioUrl, aiSet, model, eResponse, eCost) {
        let sheet = aiSheet.getSheetByName('Output');
        let timestamp = Utilities.formatDate(new Date(), aiSet.timezone, "MMMM-dd HH:mm");
        let data = [[report, output, aiSet.cost, eCost ? eCost : 'n/a', model, audioUrl, timestamp]];
        sheet.insertRowBefore(2); // Insert a new row at position 2
        sheet.getRange(2, 1, 1, data[0].length).setValues(data); // Insert the new data at the new row 2

        if (eResponse && eCost) { // expert mode output
            let expertRangeName = 'e_' + report;
            let expertRange = aiSheet.getRangeByName(expertRangeName);
            if (expertRange) {
                expertRange.setValue(eResponse);
            } else {
                Logger.log('Named range not found for expert response: ' + expertRangeName);
            }
        }
    }

    function getReportsToGenerate(s) {
        let reports = [];
        for (let key in s) {
            // Check if key is a string and starts with 'r_' and value is true
            if (typeof key === 'string' && key.startsWith('r_') && s[key] === true) {
                let reportName = key.slice(2);
                if (reportName) { // ensure we have a non-empty string after removing prefix
                    reports.push(reportName);
                }
            }
        }
        return reports;
    }

    function getDataForReport(r, u, s) {
        Logger.log('#\nGetting data for ' + r);
        // pass to cleanData the column to sort by
        switch (r) {
            case 'productTitles':
                return getDataFromSheet(s, u, 'pTitle', 'A:F', 3); // sort by cost
            case 'landingPages':
                return getDataFromSheet(s, u, 'paths', 'A:L', 1); // sort by impr & limit to 200 rows
            case 'changeHistory':
                return getDataFromSheet(s, u, 'changeData', 'B:K', -1); // limit to 200 rows, no sorting
            case 'searchCategories':
                return getDataFromSheet(s, u, 'Categories', 'C4:J', 0);
            case 'productMatrix':
                return getDataFromSheet(s, u, 'Title', 'L2:T10', 0); // Product Matrix
            case 'nGrams':
                return getDataFromSheet(s, u, 'tNgrams', 'A1:K', 1); // product title nGrams - already sorted by cost
            case 'nGramsSearch':
                return getDataFromSheet(s, u, 'sNgrams', 'A:I', 1); // search category Ngrams - already sorted by impr
            case 'asset':
                return getDataFromSheet(s, u, 'asset', 'B:P', 6); // sort by cost (col H as A not imported)
            case 'placement':
                return getDataFromSheet(s, u, 'placement', 'A:F', 2); // sort by impr 
            case 'myData':
                return getMyData();
            default:
                Logger.log('Invalid report name');
                return null;
        }
    }

    function getDataFromSheet(s, aiSheet, tabName, range, col) {
        try {
            let sheet = safeOpenAndShareSpreadsheet(aiSheet).getSheetByName(tabName);
            if (!sheet) {
                throw new Error(`Sheet ${tabName} not found in the spreadsheet.`);
            }

            let data = sheet.getRange(range).getValues().filter(row => row.some(cell => !!cell));
            if (data.length === 0 || data.length === 1) {
                throw new Error(`No data available for this report.`);
            }

            // Apply sorting and limiting logic based on col: 0 - all data; -1 limit to maxResults no sorting; else sort by col num & limit to maxResults
            if (col === 0) {
                data = data;
            } else if (col === -1) {
                data = data.slice(0, s.maxResults + 1); // account for header row
            } else {
                data = data.sort((a, b) => {
                    let valA = parseFloat(a[col]), valB = parseFloat(b[col]);
                    return valB - valA; // Sorting in descending order
                }).slice(0, s.maxResults + 1);
            }

            // Format data to string
            let formattedData = data.map(row => {
                return row.map(cell => {
                    return !isNaN(Number(cell)) ? Number(cell).toFixed(2) : cell;
                }).join(',');
            }).join('\n');
            return formattedData;
        } catch (error) {
            Logger.log(`Error in getting data from ${tabName} tab: ${error}`);
            return "No data available for this report.";
        }
    }

    function getMyData() {
        try {
            let sheet = safeOpenAndShareSpreadsheet(SHEET_URL).getSheetByName('myData');
            let lastRow = sheet.getLastRow();
            let lastColumn = sheet.getLastColumn();
            let range = sheet.getRange(1, 1, lastRow, lastColumn);
            let values = range.getValues();
            let filteredData = values.filter(row => row.some(cell => !!cell));
            let formattedData = filteredData.map(row => {
                return row.map(cell => {
                    if (!isNaN(Number(cell))) {
                        return Number(cell).toFixed(2);
                    } else {
                        return cell;
                    }
                }).join(',');
            }).join('\n');
            return formattedData;
        } catch (error) {
            Logger.log(`Error in fetching data from myData tab: ${error}`);
            return null;
        }
    }

    function textToSpeechOpenAI(text, report, aiSet) {
        Logger.log('Converting audio version of report');

        // check if no data & if that's the case, return early with message
        if (!text || text === "No data available for this report.") {
            return text;
        }

        let maxLength = 4000;
        if (text.length > maxLength) {
            Logger.log('Output is too long to use for Whisperer - proceeding with just the first 4000 characters');
        }
        text = text.length > maxLength ? text.substring(0, maxLength) : text;

        let apiUrl = 'https://api.openai.com/v1/audio/speech'; // Endpoint for OpenAI TTS

        let payload = JSON.stringify({
            model: "tts-1",
            voice: aiSet.voice,
            input: text
        });

        let options = {
            method: 'post',
            contentType: 'application/json',
            payload: payload,
            headers: {
                'Authorization': 'Bearer ' + aiSet.apiKey
            },
            muteHttpExceptions: true // To handle HTTP errors gracefully
        };

        // Make the API request
        let response = UrlFetchApp.fetch(apiUrl, options);

        let contentType = response.getHeaders()['Content-Type'];

        if (contentType.includes('audio/mpeg')) {
            try {
                // Attempt to access the folder by ID
                let f = DriveApp.getFolderById(aiSet.folder);

                // If successful, proceed with saving the file
                let blob = response.getBlob();
                // Create short timestamp for file name
                let timeStamp = Utilities.formatDate(new Date(), aiSet.timezone, "MMMdd_HHmm");
                let fileInFolder = f.createFile(blob.setName(report + ' ' + timeStamp + ".mp3"));
                let fileUrl = fileInFolder.getUrl();
                Logger.log('Audio file saved in folder: ' + fileUrl);

                return fileUrl;
            } catch (e) {
                // Handle the case where the folder does not exist
                Logger.log('Couldn\'t save audio file. Folder does not exist or access denied. Folder ID: ' + aiSet.folder);
                return null;
            }
        } else {
            // Handle unexpected content types
            Logger.log('OpenAI Text-To-Speech is having issues right now, please try again later!');
        }
    }

    function getPrompt(report, data, aiSet) {
        if (data === "No data available for this report.") {
            return data;
        }

        let promptRangeName = 'p_' + report;
        let useExpertPrompt = aiSet['use_' + report];
        let expertPromptRangeName = 'e_' + report;

        let prompt;
        try {
            let aiSheet = safeOpenAndShareSpreadsheet(aiSet.whispererUrl);

            // Get the original prompt
            let promptRange = aiSheet.getRangeByName(promptRangeName);
            if (!promptRange) {
                throw new Error(`Named range ${promptRangeName} not found`);
            }
            prompt = promptRange.getValue();

            // Check for expert prompt
            if (aiSet.expertMode && useExpertPrompt) {
                let expertPromptRange = aiSheet.getRangeByName(expertPromptRangeName);
                if (expertPromptRange) {
                    let expertPrompt = expertPromptRange.getValue();
                    if (expertPrompt && expertPrompt.trim() !== '') {
                        Logger.log(`As requested, using expert prompt for ${report}`);
                        prompt = expertPrompt;
                        // Update original prompt and clear expert prompt
                        promptRange.setValue(expertPrompt);
                        expertPromptRange.clearContent();
                        aiSheet.getRangeByName('use_' + report).setValue(false);
                        aiSet['use_' + report] = false;
                    }
                }
            }
        } catch (error) {
            Logger.log(`Error retrieving prompt for ${report}: ${error.message}`);
            prompt = "Please analyze the provided data and give insights.";
        }

        let suffix = `
        You are an expert at analyzing google ads data & providing actionable insights.
        Give all answers in the ${aiSet.lang} language. Do not use any currency symbols or currency conversion.
        Just output the numbers in the way they are. Do not round them. Provide text output, no charts.
        `;

        let usage = aiSet.whoFor === 'internal use' ? aiSet['p_internal'] : aiSet['p_client'];

        return { prompt, suffix, usage };
    }

    function logAI(ss, r, dur, aiSet) {
        let reportCost = parseFloat(aiSet.cost) + (aiSet.expertMode ? parseFloat(aiSet.expertCost) : 0);
        aiSet.totalCost += reportCost;
        let logMessage = `${r} report created in: ${dur.toFixed(0)} seconds, using ${aiSet.modelOut}, at a cost of ${aiSet.cost}`;
        logMessage += aiSet.expertMode ? ` plus ${aiSet.expertCost} for expert mode` : '';
        Logger.log(logMessage);
        let newRow = [new Date(), dur, r, aiSet.lang, aiSet.useVoice, aiSet.voice, aiSet.useEmail, aiSet.llm, aiSet.modelOut, reportCost, aiSet.ident];
        try {
            let logUrl = ss.getRangeByName('u').getValue();
            [safeOpenAndShareSpreadsheet(logUrl), ss].map(s => s.getSheetByName('log')).forEach(sheet => sheet.appendRow(newRow));
        } catch (e) {
            Logger.log('Error logging to log sheet: ' + e);
        }
    }

    function sendEmail(e, r, o, useVoice, audioUrl) {
        // check if no data present & exit early with message if that's the case
        if (o === "No data available for this report.") {
            return o;
        }
        let subject = 'Insights Created using the PMax Whisperer Script';
        let body = 'Hey<br>Your ' + r + ' insights are ready.<br>'
        if (useVoice === true && audioUrl) {
            let fileName = audioUrl.split('/').pop();
            body += '<br>Audio file: <a href="' + audioUrl + '">' + fileName + '</a>';
        }
        body += '\n' + o;

        // Sending the email
        try {
            if (e && e.includes('@')) {
                MailApp.sendEmail({
                    to: e,
                    subject: subject,
                    htmlBody: body
                });
                Logger.log('Email sent to: ' + e);
            } else {
                Logger.log('Invalid email address: ' + e);
            }
        }
        catch (e) {
            Logger.log('Error sending email: ' + e);
        }
    }

    function calculateCost(inputTokens, outputTokens, model) {
        const PRICING = {
            'gemini-2.5-flash-preview-05-20': { inputCostPerMToken: 0.15, outputCostPerMToken: 0.60 },
            'gemini-2.5-pro-preview-05-06': { inputCostPerMToken: 1.25, outputCostPerMToken: 10.00 },
            'gpt-4.1-nano-2025-04-14': { inputCostPerMToken: 0.10, outputCostPerMToken: 0.40 },
            'o4-mini-2025-04-16': { inputCostPerMToken: 1.10, outputCostPerMToken: 4.40 },
            'claude-sonnet-4-20250514': { inputCostPerMToken: 3.00, outputCostPerMToken: 15.00 },
            'claude-3-5-haiku-20241022': { inputCostPerMToken: 0.80, outputCostPerMToken: 4.00 }
        };
        // Directly access pricing for the model
        let modelPricing = PRICING[model] || { inputCostPerMToken: 1, outputCostPerMToken: 10 };
        if (!PRICING[model]) {
            Logger.log(`Default pricing of $1/m input and $10/m output used as no pricing found for model: ${model}`);
        }

        let inputCost = inputTokens * (modelPricing.inputCostPerMToken / 1e6);
        let outputCost = outputTokens * (modelPricing.outputCostPerMToken / 1e6);
        let totalCost = inputCost + outputCost;

        return totalCost.toFixed(2);
    }

    function safeOpenAndShareSpreadsheet(url, setAccess = false, newName = null) {
        try {
            // Basic validation
            if (!url) {
                console.error(`URL is empty or undefined: ${url}`);
                return null;
            }

            // Type checking and format validation
            if (typeof url !== 'string') {
                console.error(`Invalid URL type - expected string but got ${typeof url}`);
                return null;
            }

            // Validate Google Sheets URL format
            if (!url.includes('docs.google.com/spreadsheets/d/')) {
                console.error(`Invalid Google Sheets URL format: ${url}`);
                return null;
            }

            // Try to open the spreadsheet
            let ss;
            try {
                ss = SpreadsheetApp.openByUrl(url);
            } catch (error) {
                Logger.log(`Error opening spreadsheet: ${error.message}`);
                Logger.log(`Settings were: ${url}, ${setAccess}, ${newName}`);
                return null;
            }

            // Handle copy if newName is provided
            if (newName) {
                try {
                    ss = ss.copy(newName);
                } catch (error) {
                    Logger.log(`Error copying spreadsheet: ${error.message}`);
                    return null;
                }
            }

            // Handle sharing settings if required
            if (setAccess) {
                try {
                    let file = DriveApp.getFileById(ss.getId());

                    // Try ANYONE_WITH_LINK first
                    try {
                        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
                        Logger.log("Sharing set to ANYONE_WITH_LINK");
                    } catch (error) {
                        Logger.log("ANYONE_WITH_LINK failed, trying DOMAIN_WITH_LINK");

                        // If ANYONE_WITH_LINK fails, try DOMAIN_WITH_LINK
                        try {
                            file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
                            Logger.log("Sharing set to DOMAIN_WITH_LINK");
                        } catch (error) {
                            Logger.log("DOMAIN_WITH_LINK failed, setting to PRIVATE");

                            // If all else fails, set to PRIVATE
                            try {
                                file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);
                                Logger.log("Sharing set to PRIVATE");
                            } catch (error) {
                                Logger.log(`Failed to set any sharing permissions: ${error.message}`);
                            }
                        }
                    }
                } catch (error) {
                    Logger.log(`Error setting file permissions: ${error.message}`);
                    // Continue even if sharing fails - the sheet is still usable
                }
            }

            return ss;

        } catch (error) {
            // Catch any other unexpected errors
            console.error(`Unexpected error in safeOpenAndShareSpreadsheet: ${error.message}`);
            Logger.log(`Full error details: ${error.stack}`);
            return null;
        }
    }
}

function processDG(ss, s, ident, mcc, dateRangeString, elements, fromDate, toDate) {
    const dgProcessStart = new Date();
    if (!s.turnonDemandGen) {
        Logger.log('Demand Gen processing skipped as per settings.');
        return;
    }

    Logger.log('Starting Demand Gen processing.');
    let dgUrl = s.dgUrl;
    let dgSS;

    if (!dgUrl) {
        Logger.log('No DG sheet URL found. Creating one from template.');
        let dgTemplate = mcc.urls.dgTemplate;
        if (!dgTemplate) {
            Logger.log('No DG template URL in settings. Cannot create DG sheet.');
            return;
        }
        dgSS = safeOpenAndShareSpreadsheet(dgTemplate, true, `${ident} - Demand Gen Insights ${mcc.scriptVersion} - MikeRhodes.com.au (c)`);
        if (dgSS) {
            dgUrl = dgSS.getUrl();
            Logger.log(`New DG sheet created: ${dgUrl}`);
            ss.getRangeByName('dgUrl').setValue(dgUrl);
        }
    } else {
        dgSS = safeOpenAndShareSpreadsheet(dgUrl);
    }

    if (!dgSS) {
        Logger.log('Could not open or create the DG sheet. Exiting DG processing.');
        return;
    }

    try {
        dgSS.getRangeByName('pmaxUrl').setValue(ss.getUrl());
    } catch (e) {
        Logger.log('Could not set pmaxUrl in DG sheet: ' + e.message);
    }

    try {
        let range = dgSS.getRangeByName('chartDateRange');
        if (range) {
            range.setValue(dateRangeString);
        } else {
            Logger.log("Named range 'chartDateRange' not found in DG sheet. It will not be created.");
        }
    } catch (e) {
        Logger.log(`Could not set chartDateRange in DG sheet: ${e.message}`);
    }

    let queries = {};
    const e = elements;
    const date = dateRangeString;
    queries.dgCampaignQuery = 'SELECT ' + [e.segDate, e.campName, e.campId, e.networkType, e.impr, e.clicks, e.cost, e.conv, e.value, e.views, e.cpv].join(',') +
        ' FROM campaign WHERE ' + date + e.dgOnly + e.campLike + e.impr0 + e.order;
    queries.dgAdsQuery = 'SELECT ' + [e.dgAdId, e.dgAdName, e.dgAdType, e.dgFinalUrls, e.dgAdStatus, e.dgAdGroupName, e.dgAdGroupId, e.campName, e.campId, e.impr, e.clicks, e.cost, e.conv, e.value].join(',') +
        ' FROM ad_group_ad WHERE ad_group_ad.ad.type IN ("DEMAND_GEN_PRODUCT_AD", "DEMAND_GEN_VIDEO_RESPONSIVE_AD", "DEMAND_GEN_CAROUSEL_AD", "DEMAND_GEN_MULTI_ASSET_AD") AND ' + date + e.impr0;
    queries.dgAssetsQuery = 'SELECT ' + [e.dgAssetId, e.dgAssetName, e.assetSource, e.dgAssetType, e.dgImageFileSize, e.dgImageMimeType, e.dgImageUrl, e.dgImageWidth, e.dgImageHeight, e.dgTextContent, e.dgYouTubeId, e.dgYouTubeTitle, e.dgCTA, e.dgAdId, e.dgAdName, e.dgAdType, e.dgAdGroupName, e.campName, e.impr, e.clicks, e.cost, e.conv, e.value].join(',') +
        ' FROM ad_group_ad_asset_view WHERE ad_group_ad.ad.type IN ("DEMAND_GEN_PRODUCT_AD", "DEMAND_GEN_VIDEO_RESPONSIVE_AD", "DEMAND_GEN_CAROUSEL_AD", "DEMAND_GEN_MULTI_ASSET_AD") AND ' + date + e.impr0;
    queries.dgConversionsQuery = 'SELECT ' + [e.segDate, e.campName, e.campId, e.networkType, e.impr, e.clicks, e.cost, e.conv, e.value, e.views, e.cpv, e.platConv, e.platConvByDate, e.platConvIntRate, e.platConvIntValPerInt, e.platConvVal, e.platConvValByDate, e.platConvValPerCost].join(',') +
        ' FROM campaign WHERE ' + date + e.dgOnly + e.campLike + e.impr0 + e.order;
    queries.dgPlacementQuery = 'SELECT ' + [e.campName, e.placement, e.placeType, e.impr, e.inter, e.views, e.cost, e.conv, e.value].join(',') +
        ' FROM detail_placement_view WHERE ' + date + e.dgOnly + e.campLike + ' ORDER BY metrics.impressions DESC ';

    let dgCampaignData = fetchData(queries.dgCampaignQuery);
    let dgAdsData = fetchData(queries.dgAdsQuery);
    let dgAssetsData = fetchData(queries.dgAssetsQuery);
    let dgConversionsData = fetchData(queries.dgConversionsQuery);
    let dgPlacementData = fetchData(queries.dgPlacementQuery);

    // Find and set top DG campaign by cost
    let topDgCampaignName = '';
    if (dgCampaignData && dgCampaignData.length > 0) {
        const topCampaign = dgCampaignData.reduce((max, row) => {
            const maxCost = max ? (parseInt(max['metrics.costMicros']) || 0) : 0;
            const currentCost = parseInt(row['metrics.costMicros']) || 0;
            return currentCost > maxCost ? row : max;
        }, null);
        if (topCampaign) {
            topDgCampaignName = topCampaign['campaign.name'];
        }
    }

    if (topDgCampaignName) {
        try {
            const range = dgSS.getRangeByName('selectedCampaign');
            if (range) {
                range.setValue(topDgCampaignName);
                Logger.log(`Set 'selectedCampaign' in DG sheet to top campaign: ${topDgCampaignName}`);
            } else {
                Logger.log("Named range 'selectedCampaign' not found in DG sheet. It will not be created.");
            }
        } catch (e) {
            Logger.log(`Could not set 'selectedCampaign' in DG sheet: ${e.message}`);
        }
    }

    let dDgSummary = processDgSummary(dgCampaignData);
    let dDgChartData = processDgChartData(dgCampaignData);
    let dDgAds = processDgAds(dgAdsData);
    let dDgAssets = processDgAssets(dgAssetsData);
    let dAggregatedDgAssets = processAggregatedDgAssets(dgAssetsData);
    let dDgConversions = processDgConversions(dgConversionsData);
    let dDgPlacements = processOtherPlacementData(dgPlacementData, s);

    if (RUN_ASSET_FATIGUE_ANALYSIS) {
        Logger.log('Running Demand Gen Asset Fatigue Analysis...');
    let assetIds = dAggregatedDgAssets.slice(1).map(row => row[0]); // Get asset IDs, skip header
    let dAssetFatigue = processAssetFatigue(assetIds, elements, s);

    // Create a map of Asset ID -> Trend from the fatigue data
    const trendMap = new Map(dAssetFatigue.slice(1).map(row => [row[0], row[8]])); // Asset ID is at index 0, Trend is at index 8

    // Add Trend column to the aggregated assets data
    if (dAggregatedDgAssets.length > 1 && trendMap.size > 0) {
        // Add header - between CPA and AssetType
        dAggregatedDgAssets[0].splice(14, 0, 'Trend');

        // Add trend data to each row
        for (let i = 1; i < dAggregatedDgAssets.length; i++) {
            const assetId = dAggregatedDgAssets[i][0];
            const trend = trendMap.get(assetId) || 'Insufficient Data';
            dAggregatedDgAssets[i].splice(14, 0, trend);
            }
        }
        // Only output to the fatigue sheet if the analysis was run
        outputAndFormatDgSheet(dgSS, 'assetFatigue', dAssetFatigue);
    } else {
        Logger.log('Skipping Demand Gen Asset Fatigue Analysis as per script settings.');
        // Clear the sheet to avoid showing stale data and leave a message
        let fatigueSheet = dgSS.getSheetByName('assetFatigue');
        if (fatigueSheet) {
            fatigueSheet.clearContents();
            fatigueSheet.getRange('A1').setValue('Fatigue Analysis is turned off in the script settings (RUN_ASSET_FATIGUE_ANALYSIS).');
        }
    }

    outputAndFormatDgSheet(dgSS, 'dgSummary', dDgSummary);
    outputAndFormatDgSheet(dgSS, 'dgChartData', dDgChartData);
    outputAndFormatDgSheet(dgSS, 'dgAds', dDgAds);
    outputAndFormatDgSheet(dgSS, 'dgAssets', dDgAssets);
    outputAndFormatDgSheet(dgSS, 'asset', dAggregatedDgAssets); // This is still output, just without the 'Trend' column if fatigue is off
    outputAndFormatDgSheet(dgSS, 'dgConversions', dDgConversions);
    outputAndFormatDgSheet(dgSS, 'dgPlacements', dDgPlacements);

    const dgProcessTime = ((new Date() - dgProcessStart) / 1000).toFixed(2);
    if (DEBUG_LOGS_ON) Logger.log(`Demand Gen processing completed in ${dgProcessTime} seconds`);

    return { dgUrl, dgSummary: dDgSummary };

    function outputAndFormatDgSheet(ss, sheetName, data) {
        let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

        if (!data || data.length === 0) {
            Logger.log('No data to write to ' + sheetName);
            sheet.clearContents();
            return;
        }

        sheet.clearContents();

        let range = sheet.getRange(1, 1, data.length, data[0].length);
        range.setValues(data);

        const numberFormats = {
            // 0 dp
            'Impr': '#,##0',
            'Impr.': '#,##0',
            'Clicks': '#,##0',
            'Views': '#,##0',
            // 1 dp
            'ROAS': '#,##0.0',
            // 2 dp
            'Cost': '#,##0.00',
            'Value': '#,##0.00',
            'AOV': '#,##0.00',
            'Conv': '#,##0.0',
            'Conv.': '#,##0.0',
            'CPA': '#,##0.00',
            'Avg. CPV': '#,##0.00',
            'AvgCPV': '#,##0.00',
            'Camp Cost': '#,##0.00',
            'Camp Conv': '#,##0.0',
            'Camp Value': '#,##0.00',
            'Youtube Cost': '#,##0.00',
            'Youtube Conv': '#,##0.0',
            'Youtube Value': '#,##0.00',
            'Discover Cost': '#,##0.00',
            'Discover Conv': '#,##0.0',
            'Discover Value': '#,##0.00',
            'Gmail Cost': '#,##0.00',
            'Gmail Conv': '#,##0.00',
            'Gmail Value': '#,##0.00',
            // % 2dp
            'CTR': '0.00%',
            'CvR': '0.00%',
            'CVR': '0.00%',
            'YouTube Cost %': '0%',
            'YouTube Conv %': '0%',
            'YouTube Value %': '0%',
            'Discover Cost %': '0%',
            'Discover Conv %': '0%',
            'Discover Value %': '0%',
            'Gmail Cost %': '0%',
            'Gmail Conv %': '0%',
            'Gmail Value %': '0%',
            // assetFatigue formats
            'CTR Week 1': '0.00%',
            'CTR Week 2': '0.00%',
            'CTR Week 3': '0.00%',
            'CTR Week 4': '0.00%',
            'CTR Week 5': '0.00%',
            'CTR Week 6': '0.00%',
            // Platform Comparable Conversions
            'PlatConv': '#,##0.0',
            'PlatConvByDate': '#,##0.0',
            'PlatConvIntRate': '0.00%',
            'PlatConvIntValInt': '#,##0.00',
            'PlatConvVal': '#,##0.00',
            'PlatConvValByDate': '#,##0.00',
            'PlatConvValCost': '#,##0.00',
        };

        const headers = data[0];
        const rowFormats = headers.map(h => numberFormats[h] || '@');
        const formatMatrix = [];
        for (let i = 1; i < data.length; i++) {
            formatMatrix.push(rowFormats);
        }

        if (formatMatrix.length > 0) {
            sheet.getRange(2, 1, formatMatrix.length, formatMatrix[0].length).setNumberFormats(formatMatrix);
        }
    }

    function processAggregatedDgAssets(data) {
        if (!data || data.length === 0) return [['No aggregated DG Asset data available']];

        const aggregated = {};

        data.forEach(row => {
            const assetId = row['asset.id'];
            if (!assetId) return;

            if (!aggregated[assetId]) {
                let fileTitle = '';
                let urlId = '';
                switch (row['asset.type']) {
                    case 'TEXT':
                        fileTitle = row['asset.textAsset.text'];
                        urlId = assetId;
                        break;
                    case 'YOUTUBE_VIDEO':
                        fileTitle = row['asset.youtubeVideoAsset.youtubeVideoTitle'];
                        urlId = row['asset.youtubeVideoAsset.youtubeVideoId'] ? `https://www.youtube.com/watch?v=${row['asset.youtubeVideoAsset.youtubeVideoId']}` : '';
                        break;
                    case 'IMAGE':
                        fileTitle = row['asset.name'];
                        urlId = row['asset.imageAsset.fullSize.url'] || '';
                        break;
                    case 'CALL_TO_ACTION':
                        fileTitle = row['asset.callToActionAsset.callToAction'];
                        urlId = assetId;
                        break;
                }

                aggregated[assetId] = {
                    'ID': assetId,
                    'Source': row['asset.source'],
                    'File/Title': fileTitle,
                    'URL/ID': urlId,
                    'Asset Type': row['asset.type'],
                    'Impr': 0, 'Clicks': 0, 'Cost': 0, 'Conv': 0, 'Value': 0
                };
            }

            aggregated[assetId]['Impr'] += parseInt(row['metrics.impressions']) || 0;
            aggregated[assetId]['Clicks'] += parseInt(row['metrics.clicks']) || 0;
            aggregated[assetId]['Cost'] += (parseInt(row['metrics.costMicros']) || 0) / 1e6;
            aggregated[assetId]['Conv'] += parseFloat(row['metrics.conversions']) || 0;
            aggregated[assetId]['Value'] += parseFloat(row['metrics.conversionsValue']) || 0;
        });

        const headers = ['ID', 'Source', 'File/Title', 'URL/ID', 'Impr', 'Clicks', 'Cost', 'Conv', 'Value', 'CTR', 'CVR', 'AOV', 'ROAS', 'CPA', 'AssetType'];

        const rows = Object.values(aggregated).map(agg => {
            const impr = agg['Impr'];
            const clicks = agg['Clicks'];
            const cost = agg['Cost'];
            const conv = agg['Conv'];
            const value = agg['Value'];

            const ctr = impr > 0 ? clicks / impr : 0;
            const cvr = clicks > 0 ? conv / clicks : 0;
            const aov = conv > 0 ? value / conv : 0;
            const roas = cost > 0 ? value / cost : 0;
            const cpa = conv > 0 ? cost / conv : 0;

            return [
                agg['ID'], agg['Source'], agg['File/Title'], agg['URL/ID'],
                impr, clicks, cost, conv, value,
                ctr, cvr, aov, roas, cpa,
                agg['Asset Type']
            ];
        });

        rows.sort((a, b) => b[4] - a[4]); // Sort by Impr descending

        return [headers, ...rows];
    }

    function processNetworkBreakdown(data, groupByDate) {
        if (!data || data.length === 0) return [['No DG network breakdown data available']];
        const aggregated = {};
        data.forEach(row => {
            const campaignName = row['campaign.name'];
            const date = row['segments.date'];
            const key = groupByDate ? `${date}_${campaignName}` : campaignName;

            if (!aggregated[key]) {
                aggregated[key] = {
                    'Date': groupByDate ? date : '',
                    'Campaign Name': campaignName,
                    'Camp Cost': 0, 'Camp Conv': 0, 'Camp Value': 0,
                    'Youtube Cost': 0, 'Youtube Conv': 0, 'Youtube Value': 0,
                    'Discover Cost': 0, 'Discover Conv': 0, 'Discover Value': 0,
                    'Gmail Cost': 0, 'Gmail Conv': 0, 'Gmail Value': 0
                };
            }

            const cost = (row['metrics.costMicros'] || 0) / 1e6;
            const conv = parseFloat(row['metrics.conversions']) || 0;
            const value = parseFloat(row['metrics.conversionsValue']) || 0;
            const network = row['segments.adNetworkType'];

            aggregated[key]['Camp Cost'] += cost;
            aggregated[key]['Camp Conv'] += conv;
            aggregated[key]['Camp Value'] += value;

            if (network === 'YOUTUBE') {
                aggregated[key]['Youtube Cost'] += cost;
                aggregated[key]['Youtube Conv'] += conv;
                aggregated[key]['Youtube Value'] += value;
            } else if (network === 'DISCOVER') {
                aggregated[key]['Discover Cost'] += cost;
                aggregated[key]['Discover Conv'] += conv;
                aggregated[key]['Discover Value'] += value;
            } else if (network === 'GMAIL') {
                aggregated[key]['Gmail Cost'] += cost;
                aggregated[key]['Gmail Conv'] += conv;
                aggregated[key]['Gmail Value'] += value;
            }
        });

        const headers = [
            'Campaign Name', 'Camp Cost', 'Camp Conv', 'Camp Value',
            'Youtube Cost', 'Youtube Conv', 'Youtube Value',
            'Discover Cost', 'Discover Conv', 'Discover Value',
            'Gmail Cost', 'Gmail Conv', 'Gmail Value',
            'YouTube Cost %', 'Discover Cost %', 'Gmail Cost %',
            'YouTube Conv %', 'Discover Conv %', 'Gmail Conv %',
            'YouTube Value %', 'Discover Value %', 'Gmail Value %',
            'Youtube Cost', 'Discover Cost', 'Gmail Cost',
            'Youtube Conv', 'Discover Conv', 'Gmail Conv',
            'Youtube Value', 'Discover Value', 'Gmail Value'
        ];

        if (groupByDate) {
            headers.unshift('Date');
        }

        const rows = Object.values(aggregated).map(agg => {
            const campCost = agg['Camp Cost'];
            const campConv = agg['Camp Conv'];
            const campValue = agg['Camp Value'];
            let row = [
                agg['Campaign Name'],
                agg['Camp Cost'], agg['Camp Conv'], agg['Camp Value'],
                // By Network type
                agg['Youtube Cost'], agg['Youtube Conv'], agg['Youtube Value'],
                agg['Discover Cost'], agg['Discover Conv'], agg['Discover Value'],
                agg['Gmail Cost'], agg['Gmail Conv'], agg['Gmail Value'],
                // Cost Percentages
                campCost > 0 ? agg['Youtube Cost'] / campCost : 0,
                campCost > 0 ? agg['Discover Cost'] / campCost : 0,
                campCost > 0 ? agg['Gmail Cost'] / campCost : 0,
                // Conversion Percentages
                campConv > 0 ? agg['Youtube Conv'] / campConv : 0,
                campConv > 0 ? agg['Discover Conv'] / campConv : 0,
                campConv > 0 ? agg['Gmail Conv'] / campConv : 0,
                // Value Percentages
                campValue > 0 ? agg['Youtube Value'] / campValue : 0,
                campValue > 0 ? agg['Discover Value'] / campValue : 0,
                campValue > 0 ? agg['Gmail Value'] / campValue : 0,
                // Repeat for totals in different grouping to make it easy to chart
                agg['Youtube Cost'], agg['Discover Cost'], agg['Gmail Cost'],
                agg['Youtube Conv'], agg['Discover Conv'], agg['Gmail Conv'],
                agg['Youtube Value'], agg['Discover Value'], agg['Gmail Value']
            ];
            if (groupByDate) {
                row.unshift(agg['Date']);
            }
            return row;
        });

        if (groupByDate) {
            // dgSummary: Sort by Date asc, then Camp Cost desc
            rows.sort((a, b) => {
                if (a[0] > b[0]) return 1;
                if (a[0] < b[0]) return -1;
                return b[2] - a[2];
            });
        } else {
            // dgChartData: Sort by Camp Cost desc
            rows.sort((a, b) => b[1] - a[1]);
        }

        return [headers, ...rows];
    }

    function processDgSummary(data) {
        return processNetworkBreakdown(data, true);
    }

    function processDgChartData(data) {
        return processNetworkBreakdown(data, false);
    }

    function processDgAds(data) {
        if (!data || data.length === 0) return [['No Demand Gen Ads data available']];
        const headers = ['Campaign', 'AdGroup', 'AdName', 'AdID', 'AdType', 'Status', 'Impr', 'Clicks', 'Cost', 'Conv.', 'Value', 'CTR', 'CvR', 'CPA', 'ROAS'];
        const rows = data.map(row => {
            const cost = (row['metrics.costMicros'] || 0) / 1e6;
            const clicks = row['metrics.clicks'] || 0;
            const impressions = row['metrics.impressions'] || 0;
            const conversions = row['metrics.conversions'] || 0;
            const value = row['metrics.conversionsValue'] || 0;
            return [
                row['campaign.name'],
                row['adGroup.name'],
                row['adGroupAd.ad.name'],
                row['adGroupAd.ad.id'],
                row['adGroupAd.ad.type'],
                row['adGroupAd.status'],
                impressions,
                clicks,
                cost,
                conversions,
                value,
                impressions > 0 ? clicks / impressions : 0,
                clicks > 0 ? conversions / clicks : 0,
                conversions > 0 ? cost / conversions : 0,
                cost > 0 ? value / cost : 0
            ];
        });
        return [headers, ...rows];
    }

    function processDgAssets(data) {
        if (!data || data.length === 0) return [['No Demand Gen Asset data available']];
        const headers = ['Campaign', 'AdGroup', 'AdName', 'AdID', 'AssetName', 'AssetID', 'AssetType', 'Content', 'Impr', 'Clicks', 'Cost', 'Conv.', 'Value', 'CTR', 'CvR', 'CPA', 'ROAS'];
        const rows = data.map(row => {
            const cost = (row['metrics.costMicros'] || 0) / 1e6;
            const clicks = row['metrics.clicks'] || 0;
            const impressions = row['metrics.impressions'] || 0;
            const conversions = row['metrics.conversions'] || 0;
            const value = row['metrics.conversionsValue'] || 0;
            let content = '';
            switch (row['asset.type']) {
                case 'TEXT': content = row['asset.textAsset.text']; break;
                case 'YOUTUBE_VIDEO': content = row['asset.youtubeVideoAsset.youtubeVideoId'] ? `https://www.youtube.com/watch?v=${row['asset.youtubeVideoAsset.youtubeVideoId']}` : ''; break;
                case 'IMAGE': content = row['asset.imageAsset.fullSize.url'] || row['asset.name']; break;
                case 'CALL_TO_ACTION': content = row['asset.callToActionAsset.callToAction']; break;
            }
            return [
                row['campaign.name'],
                row['adGroup.name'],
                row['adGroupAd.ad.name'],
                row['adGroupAd.ad.id'],
                row['asset.name'],
                row['asset.id'],
                row['asset.type'],
                content,
                impressions,
                clicks,
                cost,
                conversions,
                value,
                impressions > 0 ? clicks / impressions : 0,
                clicks > 0 ? conversions / clicks : 0,
                conversions > 0 ? cost / conversions : 0,
                cost > 0 ? value / cost : 0
            ];
        });
        return [headers, ...rows];
    }

    function processAssetFatigue(assetIds, e, s) {
        if (!assetIds || assetIds.length === 0) {
            return [['No assets to analyze for fatigue.']];
        }

        // 1. Calculate date range for the last 6 full weeks (Mon-Sun)
        let today = new Date();
        let dayOfWeek = today.getDay(); // Sunday = 0, Monday = 1, etc.
        let daysToLastSunday = dayOfWeek === 0 ? 7 : dayOfWeek;
        let endDate = new Date(today);
        endDate.setDate(today.getDate() - daysToLastSunday);

        let startDate = new Date(endDate);
        startDate.setDate(endDate.getDate() - 41); // 6 weeks total (6*7 - 1)

        const fStartDate = Utilities.formatDate(startDate, s.timezone, 'yyyy-MM-dd');
        const fEndDate = Utilities.formatDate(endDate, s.timezone, 'yyyy-MM-dd');
        const dateRange = `segments.date BETWEEN '${fStartDate}' AND '${fEndDate}'`;

        // 2. Build and run the query
        // Chunk asset IDs to avoid query length limits
        const chunkSize = 1000;
        let allWeeklyData = [];
        for (let i = 0; i < assetIds.length; i += chunkSize) {
            const assetIdChunk = assetIds.slice(i, i + chunkSize);
            const assetIdFilter = `asset.id IN (${assetIdChunk.map(id => `'${id}'`).join(',')})`;
            const weeklyQuery = `
          SELECT asset.id, asset.name, asset.youtube_video_asset.youtube_video_title, asset.text_asset.text, segments.week, metrics.clicks, metrics.impressions
          FROM ad_group_ad_asset_view
          WHERE ${assetIdFilter} AND ${dateRange} AND ad_group_ad.ad.type IN ("DEMAND_GEN_PRODUCT_AD", "DEMAND_GEN_VIDEO_RESPONSIVE_AD", "DEMAND_GEN_CAROUSEL_AD", "DEMAND_GEN_MULTI_ASSET_AD")`;
            let weeklyData = fetchData(weeklyQuery);
            allWeeklyData = allWeeklyData.concat(weeklyData);
        }


        if (allWeeklyData.length === 0) {
            return [['No weekly performance data found for these assets in the last 6 weeks.']];
        }

        // 3. Process the data
        const assetData = {};
        allWeeklyData.forEach(row => {
            const id = row['asset.id'];
            if (!assetData[id]) {
                assetData[id] = {
                    title: row['asset.name'] || row['asset.youtubeVideoAsset.youtubeVideoTitle'] || row['asset.textAsset.text'] || `Asset ${id}`,
                    weeks: {}
                };
            }
            assetData[id].weeks[row['segments.week']] = {
                clicks: parseInt(row['metrics.clicks']) || 0,
                impressions: parseInt(row['metrics.impressions']) || 0,
            };
        });

        // 4. Get the last 6 week start dates for headers
        const weekDates = [];
        let lastMonday = new Date(endDate);
        lastMonday.setDate(endDate.getDate() - 6); // Find the Monday of the last full week
        for (let i = 5; i >= 0; i--) {
            let weekStart = new Date(lastMonday);
            weekStart.setDate(lastMonday.getDate() - (i * 7));
            weekDates.push(Utilities.formatDate(weekStart, 'UTC', 'yyyy-MM-dd'));
        }

        const headers = ['Asset ID', 'File-Title', 'CTRWeek1', 'CTRWeek2', 'CTRWeek3', 'CTRWeek4', 'CTRWeek5', 'CTRWeek6', 'Trend'];
        const results = [];

        // 5. Analyze each asset
        for (const id in assetData) {
            const asset = assetData[id];
            const ctrs = weekDates.map(week => {
                const weekData = asset.weeks[week];
                if (weekData && weekData.impressions > 0) {
                    return weekData.clicks / weekData.impressions;
                }
                return null; // Use null for weeks with no data
            });

            const validCtrs = ctrs.filter(ctr => ctr !== null);
            let trend = 'Insufficient Data';

            if (validCtrs.length >= 3) { // Need at least 3 data points for a trend
                const trendAnalysis = linearRegression(validCtrs.map((ctr, i) => ({ x: i, y: ctr })));
                trend = getTrendCategory(trendAnalysis.slope, trendAnalysis.r2);
            }

            const formattedCtrs = ctrs.map(ctr => ctr === null ? 'N/A' : ctr);
            results.push([id, asset.title, ...formattedCtrs, trend]);
        }

        return [headers, ...results];
    }

    function processOtherPlacementData(data, s) {
        if (!data || data.length === 0) {
            Logger.log('No placement data available for non-PMax campaigns.');
            return [['No data available']];
        }

        let minPlaceImpr = s.minPlaceImpr;

        const aggregatedData = new Map();
        data.forEach(row => {
            const key = row['detailPlacementView.groupPlacementTargetUrl'];
            if (!aggregatedData.has(key)) {
                aggregatedData.set(key, {
                    placement: key,
                    type: row['detailPlacementView.placementType'],
                    impressions: 0, interactions: 0, views: 0, cost: 0, conversions: 0, value: 0
                });
            }
            const entry = aggregatedData.get(key);
            entry.impressions += Number(row['metrics.impressions']) || 0;
            entry.interactions += Number(row['metrics.interactions']) || 0;
            entry.views += Number(row['metrics.videoViews']) || 0;
            entry.cost += (Number(row['metrics.costMicros']) || 0) / 1e6;
            entry.conversions += Number(row['metrics.conversions']) || 0;
            entry.value += Number(row['metrics.conversionsValue']) || 0;
        });

        let processedData = [];
        const headers = [['Campaign', 'Placement', 'Type', 'Impr.', 'Interactions', 'Views', 'Cost', 'Conv.', 'Value', 'CTR', 'CVR', 'CPA', 'ROAS']];

        for (const placementData of aggregatedData.values()) {
            if (placementData.impressions >= minPlaceImpr) {
                const cost = placementData.cost;
                const interactions = placementData.interactions;
                const impressions = placementData.impressions;
                const conversions = placementData.conversions;
                const value = placementData.value;
                processedData.push([
                    'All Campaigns',
                    placementData.placement,
                    placementData.type,
                    impressions,
                    interactions,
                    placementData.views,
                    cost,
                    conversions,
                    value,
                    impressions > 0 ? interactions / impressions : 0,
                    interactions > 0 ? conversions / interactions : 0,
                    conversions > 0 ? cost / conversions : 0,
                    cost > 0 ? value / cost : 0
                ]);
            }
        }

        processedData.sort((a, b) => b[3] - a[3]);
        const aggregatedLength = processedData.length;

        data.forEach(row => {
            const impressions = parseInt(row['metrics.impressions']) || 0;
            if (impressions >= minPlaceImpr) {
                const cost = (Number(row['metrics.costMicros']) || 0) / 1e6;
                const interactions = Number(row['metrics.interactions']) || 0;
                const conversions = Number(row['metrics.conversions']) || 0;
                const value = Number(row['metrics.conversionsValue']) || 0;
                processedData.push([
                    row['campaign.name'],
                    row['detailPlacementView.groupPlacementTargetUrl'],
                    row['detailPlacementView.placementType'],
                    impressions,
                    interactions,
                    Number(row['metrics.videoViews']) || 0,
                    cost,
                    conversions,
                    value,
                    impressions > 0 ? interactions / impressions : 0,
                    interactions > 0 ? conversions / interactions : 0,
                    conversions > 0 ? cost / conversions : 0,
                    cost > 0 ? value / cost : 0
                ]);
            }
        });

        const campaignData = processedData.slice(aggregatedLength);
        campaignData.sort((a, b) => {
            if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
            return b[3] - a[3];
        });

        return [...headers, ...processedData.slice(0, aggregatedLength), ...campaignData];
    }

    function linearRegression(data) {
        const n = data.length;
        let sum_x = 0, sum_y = 0, sum_xy = 0, sum_xx = 0;

        for (let i = 0; i < n; i++) {
            sum_x += data[i].x;
            sum_y += data[i].y;
            sum_xy += (data[i].x * data[i].y);
            sum_xx += (data[i].x * data[i].x);
        }

        const slope = (n * sum_xy - sum_x * sum_y) / (n * sum_xx - sum_x * sum_x);
        const intercept = (sum_y - slope * sum_x) / n;

        // Calculate R-squared
        let ss_tot = 0, ss_res = 0;
        const y_mean = sum_y / n;
        for (let i = 0; i < n; i++) {
            ss_tot += Math.pow(data[i].y - y_mean, 2);
            ss_res += Math.pow(data[i].y - (slope * data[i].x + intercept), 2);
        }
        const r2 = (ss_tot === 0) ? 1 : 1 - (ss_res / ss_tot);

        return { slope, intercept, r2 };
    }

    function getTrendCategory(slope, r2) {
        const R2_THRESHOLD_HIGH = 0.65;
        const R2_THRESHOLD_LOW = 0.3;
        const SLOPE_THRESHOLD_HIGH = 0.002; // 0.2% CTR change per week
        const SLOPE_THRESHOLD_LOW = 0.0005; // 0.05% CTR change per week

        if (Math.abs(slope) < SLOPE_THRESHOLD_LOW) return 'Sideways';

        const direction = slope > 0 ? 'Up' : 'Down';
        const magnitude = Math.abs(slope) > SLOPE_THRESHOLD_HIGH ? 'Clear' : 'Slight';

        if (r2 > R2_THRESHOLD_HIGH) {
            return `${magnitude} Trend ${direction}`;
        } else if (r2 < R2_THRESHOLD_LOW) {
            return 'Volatile';
        } else {
            return `Potential Trend ${direction}`;
        }
    }

    function processDgConversions(data) {
        if (!data || data.length === 0) return [['No DG Conversion data available']];
        const headers = [
            'Date', 'Campaign', 'Campaign ID', 'Network', 'Impr', 'Clicks', 'Cost', 'Conv', 'Value', 'Views', 'Avg. CPV',
            'PlatConv', 'PlatConvByDate', 'PlatConvIntRate', 'PlatConvIntValInt',
            'PlatConvVal', 'PlatConvValByDate', 'PlatConvValCost'
        ];
        const rows = data.map(row => {
            return [
                row['segments.date'],
                row['campaign.name'],
                row['campaign.id'],
                row['segments.adNetworkType'],
                row['metrics.impressions'] || 0,
                row['metrics.clicks'] || 0,
                (row['metrics.costMicros'] || 0) / 1e6,
                row['metrics.conversions'] || 0,
                row['metrics.conversionsValue'] || 0,
                row['metrics.videoViews'] || 0,
                (row['metrics.averageCpv'] || 0) / 1e6,
                row['metrics.platformComparableConversions'] || 0,
                row['metrics.platformComparableConversionsByConversionDate'] || 0,
                row['metrics.platformComparableConversionsFromInteractionsRate'] || 0,
                row['metrics.platformComparableConversionsFromInteractionsValuePerInteraction'] || 0,
                row['metrics.platformComparableConversionsValue'] || 0,
                row['metrics.platformComparableConversionsValueByConversionDate'] || 0,
                row['metrics.platformComparableConversionsValuePerCost'] || 0
            ];
        });
        return [headers, ...rows];
    }
}
//#endregion

//#region rest of pmax script -------------------------------------------------------------------------
function processAdvancedSettings(ss, s, q, mainDateRange, e) {
    // For pTitleCampaign, pTitleID, paths handle preserving the last column
    const clearSheetExceptLastColumn = (sheetName) => {
        const sheet = ss.getSheetByName(sheetName);
        const lastColumn = sheet.getMaxColumns();
        const rangeToClear = sheet.getRange(2, 1, sheet.getMaxRows() - 1, lastColumn - 1);
        rangeToClear.clearContent();
    };

    // For other tabs - General function to clear data
    const clearSheet = (sheetName) => {
        const sheet = ss.getSheetByName(sheetName);
        sheet.clearContents(); // This clears the entire sheet, but not formats and values
    };

    // Handling turnonTitleCampaign
    if (s.turnonTitleCampaign) {
        let titleCampaignData = fetchProductData(q.productQuery, s.tCost, s.tRoas, 'pTitleCampaign');
        outputDataToSheet(ss, 'pTitleCampaign', titleCampaignData, 'notLast');
        ss.getSheetByName('TitleCampaign').showSheet();
    } else {
        clearSheetExceptLastColumn('pTitleCampaign');
        ss.getSheetByName('pTitleCampaign').hideSheet();
    }

    // Handling turnonTitleID
    if (s.turnonTitleID) {
        let titleIDData = fetchProductData(q.productQuery, s.tCost, s.tRoas, 'pTitleID');
        outputDataToSheet(ss, 'pTitleID', titleIDData, 'notLast');
        ss.getSheetByName('Title&ID').showSheet();
    } else {
        clearSheetExceptLastColumn('pTitleID');
        ss.getSheetByName('pTitleID').hideSheet();
    }

    // Handling turnonIDChannel
    if (s.turnonIDChannel) {
        let idChannelData = fetchProductData(q.productQuery, s.tCost, s.tRoas, 'idChannel');
        outputDataToSheet(ss, 'idChannel', idChannelData, 'clear');
        ss.getSheetByName('idChannel').showSheet();
    } else {
        clearSheet('idChannel');
        ss.getSheetByName('idChannel').hideSheet();
    }


    // Handling turnonGeo
    if (s.turnonGeo) {
        Logger.log('Starting geo field discovery...');

        const availableGeoFields = discoverAvailableGeoFields(e, mainDateRange);

        // Build dynamic query with all available geo fields
        const allGeoFields = [e.geoCountryId, ...availableGeoFields.map(f => f.field)];
        const dynamicGeoQuery = buildGeoPerformanceQuery(e, allGeoFields, mainDateRange);

        let geoPerformanceData = fetchData(dynamicGeoQuery);
        let locationData = fetchData(q.locationDataQuery);

        if (geoPerformanceData && geoPerformanceData.length > 0) {
            let geo = processGeoPerformanceData(geoPerformanceData, locationData, ss, availableGeoFields);
            outputDataToSheet(ss, 'geo_locations', geo.locations, 'clear');
            outputDataToSheet(ss, 'geo_campaigns', geo.campaigns, 'clear');
        } else {
            Logger.log('No geo performance data found for the selected date range.');
            outputDataToSheet(ss, 'geo_locations', [['No geo performance data found.']], 'clear');
            outputDataToSheet(ss, 'geo_campaigns', [['No geo performance data found.']], 'clear');
        }
        ss.getSheetByName('Geo').showSheet();
    } else {
        clearSheet('geo_locations');
        clearSheet('geo_campaigns');
        ss.getSheetByName('geo_locations').hideSheet();
        ss.getSheetByName('geo_campaigns').hideSheet();
        ss.getSheetByName('Geo').hideSheet();
    }

    // Handling turnonChange
    if (s.turnonChange) {
        let changeData = fetchData(q.changeQuery);
        outputDataToSheet(ss, 'changeData', changeData, 'clear');
        ss.getSheetByName('changeData').showSheet();
    } else {
        clearSheet('changeData');
        ss.getSheetByName('changeData').hideSheet();
    }

    // Handling turnonLP - landing page data - keep last column (flag for Pages)
    if (s.turnonLP) {
        let lpData = fetchData(q.lpQuery);
        aggLPData(lpData, ss, s, 'paths');
        ss.getSheetByName('Pages').showSheet();
    } else {
        try {
            clearSheetExceptLastColumn('paths');
            ss.getSheetByName('Pages').hideSheet();
            ss.getSheetByName('paths').hideSheet();
        } catch (error) {
            Logger.log('Error: The paths tab is not in your sheet. Please update to the latest template. You need v70 or later. ' + error);
        }
    }

    if (s.turnonPlace) {
        let pmaxPlacementData = fetchData(q.pmaxPlacementQuery);
        outputDataToSheet(ss, 'placement', processPmaxPlacementData(pmaxPlacementData, s), 'clear');
        ss.getSheetByName('placement').showSheet();

    } else {
        clearSheet('placement');
        ss.getSheetByName('placement').hideSheet();
    }
}

function configureScript(sheetUrl, clientCode, clientSettings, mcc) {
    let start = new Date();
    let s, ss;
    let accountName = AdsApp.currentAccount().getName();
    let ident = clientCode || accountName;
    if (mcc.comingFrom === 'single') {
        let accDate = new Date().toLocaleString('en-US', { timeZone: mcc.mccTimezone });
        mcc.mccDay = new Date(accDate).getDay();
    }

    ss = handleNormalMode(mcc.urls.clientNew || sheetUrl, ident, mcc);

    // Update AI run at & account type. Create default settings and update variables
    updateRunAI(ss, clientSettings.aiRunAt, mcc.comingFrom);
    updateAccountType(ss, clientSettings.accountType);
    updateRunDG(ss, clientSettings.runDG);
    let defaultSettings = createDefaultSettings(clientSettings, mcc);
    s = updateVariablesFromSheet(ss, mcc.scriptVersion, defaultSettings);
    s.ident = ident;
    s.minLpImpr = mcc.minLpImpr;

    // Log settings and check version
    logSettings(s, ss);
    checkVersion(mcc.scriptVersion, s.sheetVersion, ss, mcc.urls.template);

    return { ss, s, start, ident };
}

function updateRunAI(ss, aiRunAt, comingFrom) {
    if (ss && aiRunAt && comingFrom === 'mcc') {
        let convertedAiRunAt = aiRunAt;
        if (typeof aiRunAt === 'number') {
            const dayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
            if (aiRunAt === -1) {
                convertedAiRunAt = 'next script run';
            } else if (aiRunAt === 99) {
                convertedAiRunAt = 'never';
            } else if (aiRunAt >= 0 && aiRunAt <= 6) {
                convertedAiRunAt = dayNames[aiRunAt];
            }
        }
        aiRunAt = convertedAiRunAt;
        try {
            let aiRunAtRange = ss.getRangeByName('aiRunAt');
            if (aiRunAtRange) {
            aiRunAtRange.setValue(aiRunAt);
            Logger.log('Update AI sheet day: MCC setting for this client is ' + aiRunAt + ', so setting client sheet to ' + aiRunAt);
            } else {
                Logger.log('Error: aiRunAt named range not found in sheet' + error);
            }
        } catch (error) {
            Logger.log('Error: Unable to update the aiRunAt range. ' + error);
        }
    }
}

function updateAccountType(ss, accountType) {
    if (ss && accountType === 'leadgen') {
        try {
            let accountTypeRange = ss.getRange('accountType'); // This should be getRangeByName
            if (accountTypeRange) { // Add null check
            accountTypeRange.setValue('leadgen');
                Logger.log('Account type in MCC was leadgen, setting client sheet to leadgen');
            } else {
                Logger.log('Error: accountType named range not found in sheet');
            }
        } catch (error) {
            Logger.log('Error: Unable to update the accountType range. ' + error);
        }
    }
}

function updateRunDG(ss, runDG) {

    if (ss && runDG !== undefined) {
        try {
            let runDGRange = ss.getRange('turnonDemandGen'); // named range in single sheet
            runDGRange.setValue(runDG);
            Logger.log(`Run DG in MCC was ${runDG}, setting client sheet to ${runDG}`); // default is true
        } catch (error) {
            Logger.log('Error: Unable to update the runDG range. ' + error);
        }
    }
}

function handleNormalMode(sheetUrl, ident, mcc) {
    let ss;
    if (sheetUrl) {
        ss = safeOpenAndShareSpreadsheet(sheetUrl);
    } else {
        // Create the main PMax sheet
        ss = safeOpenAndShareSpreadsheet(mcc.urls.template, true, `${ident} ${mcc.scriptVersion} - PMax Insights - (c) MikeRhodes.com.au `);
        Logger.log('****\nCreated new PMax Insights sheet for: ' + ident + '\nURL is ' + ss.getUrl() +
            '\nRemember to add this URL to the top of your script before next run.\n****');

        // Since this is a new sheet, we know the DG and AI sheets also need to be created.

        // Create the DG sheet
        try {
            const dgTemplate = mcc.urls.dgTemplate;
            if (dgTemplate) {
                const dgSS = safeOpenAndShareSpreadsheet(dgTemplate, true, `${ident} - Demand Gen Insights ${mcc.scriptVersion} - MikeRhodes.com.au (c)`);
                if (dgSS) {
                    const dgUrl = dgSS.getUrl();
                    Logger.log(`New DG sheet created automatically: ${dgUrl}`);
                    ss.getRangeByName('dgUrl').setValue(dgUrl);
                    dgSS.getRangeByName('pmaxUrl').setValue(ss.getUrl()); // also set the back-reference
                }
            } else {
                Logger.log('No DG template URL in settings. Cannot create DG sheet automatically.');
            }
        } catch (e) {
            Logger.log(`Could not create DG sheet automatically. You may need to enable it on the sheet and run again. Error: ${e.message}`);
        }

        // Create the AI sheet
        try {
            const aiTemplate = mcc.urls.aiTemplate;
            if (aiTemplate) {
                const aiSheet = safeOpenAndShareSpreadsheet(aiTemplate, true, `${ident} - AI Whisperer ${mcc.scriptVersion} - MikeRhodes.com.au (c)`);
                if (aiSheet) {
                    const whispererUrl = aiSheet.getUrl();
                    Logger.log(`New AI Whisperer sheet created automatically: ${whispererUrl}`);
                    ss.getRangeByName('whispererUrl').setValue(whispererUrl);
                    aiSheet.getRangeByName('pmaxUrl').setValue(ss.getUrl()); // also set the back-reference
                }
            } else {
                Logger.log('No AI template URL in settings. Cannot create AI sheet automatically.');
            }
        } catch (e) {
            Logger.log(`Could not create AI sheet automatically. You may need to enable it on the sheet and run again. Error: ${e.message}`);
        }
    }
    return ss;
}

function createDefaultSettings(clientSettings, mcc) {
    return {
        numberOfDays: mcc.defaultSettings.numberOfDays,
        tCost: mcc.defaultSettings.tCost,
        tRoas: mcc.defaultSettings.tRoas,
        brandTerm: clientSettings.brandTerm,
        accountType: clientSettings.accountType,
        aiTemplate: mcc.urls.aiTemplate,
        dgTemplate: mcc.urls.dgTemplate,
        whispererUrl: clientSettings.whispererUrl,
        dgUrl: clientSettings.dgUrl,

        // Additional settings from 'Settings' and 'Advanced' sheets
        aiRunAt: 1, // default to monday
        fromDate: undefined,
        toDate: undefined,
        lotsProducts: 0,
        campFilter: '',
        turnonLP: false,
        turnonPlace: true,
        turnonGeo: false,
        turnonChange: false,
        turnonAISheet: true,
        turnonTitleCampaign: false,
        turnonTitleID: false,
        turnonIDChannel: false,
        turnonDemandGen: false,
        sheetVersion: '',
        scriptVersion: mcc.scriptVersion,
        levenshtein: 2
    };
}

function logSettings(s, ss) {
    s = s || {};
    let logSettings = { ...s };
    delete logSettings.levenshtein;
    delete logSettings.inputTokens;
    delete logSettings.outputTokens;
    delete logSettings.themes;
    logSettings.apiKey = s.apiKey ? s.apiKey.slice(0, 10) + '...' : '';
    logSettings.anth_apikey = s.anth_apikey ? s.anth_apikey.slice(0, 10) + '...' : '';
    logSettings.aiTemplate = s.aiTemplate && s.aiTemplate.includes('/d/') ? s.aiTemplate.split('/d/')[1].slice(0, 5) + '...' : '';
    logSettings.dgTemplate = s.dgTemplate && s.dgTemplate.includes('/d/') ? s.dgTemplate.split('/d/')[1].slice(0, 5) + '...' : '';
    logSettings.whispererUrl = s.whispererUrl && s.whispererUrl.includes('/d/') ? s.whispererUrl.split('/d/')[1].slice(0, 5) + '...' : '';
    logSettings.dgUrl = s.dgUrl && s.dgUrl.includes('/d/') ? s.dgUrl.split('/d/')[1].slice(0, 5) + '...' : '';
    logSettings.ssUrl = ss.getUrl();
    Logger.log('Settings: ' + JSON.stringify(logSettings));
}

function updateVariablesFromSheet(ss, version, defaultSettings) {

    let updatedSettings = { ...defaultSettings };

    const allRangeNames = [
        'numberOfDays', 'tCost', 'tRoas', 'brandTerm', 'accountType', 'aiRunAt', 'minPlaceImpr',
        'fromDate', 'toDate', 'lotsProducts', 'turnonTitleID', 'turnonIDChannel', 'turnonTitleCampaign', 'turnonDemandGen',
        'campFilter', 'turnonLP', 'turnonPlace', 'turnonGeo', 'turnonChange', 'turnonAISheet', 'sheetVersion', 'whispererUrl', 'dgUrl'
    ];

    try {
        allRangeNames.forEach(rangeName => {
            try {
                const range = ss.getRangeByName(rangeName);
                if (range) {
                    let value = range.getDisplayValue();
                    if (value !== null && value !== undefined) {
                        value = String(value).trim(); // Convert to string first, then trim
                    if (value !== '') {
                        if (['fromDate', 'toDate'].includes(rangeName)) {
                                updatedSettings[rangeName] = /^\d{2}\/\d{2}\/\d{4}$/.test(value) ? value : updatedSettings[rangeName];
                        } else if (rangeName.startsWith('turnon')) {
                            updatedSettings[rangeName] = value.toLowerCase() === 'true';
                        } else if (['numberOfDays', 'tCost', 'tRoas', 'lotsProducts'].includes(rangeName)) {
                            updatedSettings[rangeName] = isNaN(value) ? updatedSettings[rangeName] : Number(value);
                        } else {
                            updatedSettings[rangeName] = value;
                        }
                    }
                    } else {
                        console.warn(`Named range ${rangeName} returned null/undefined value. Using default.`);
                    }
                } else {
                    console.warn(`Named range ${rangeName} not found. Using default value.`);
                }
            } catch (e) {
                console.warn(`Named range ${rangeName} not found or error occurred: ${e.message}. Using default value.`);
            }
        });
        updatedSettings.aiRunAt = convertAIRunAt(updatedSettings.aiRunAt);
        return updatedSettings;

    } catch (e) {
        console.error("Error in 'updateVariablesFromSheet': ", e);
        return updatedSettings;
    }
}

function convertAIRunAt(aiRunAt) {
    const dayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];

    if (typeof aiRunAt == 'string' && aiRunAt) {
        aiRunAt = aiRunAt.toLowerCase();
        if (aiRunAt === 'next script run') {
            return -1;
        } else if (aiRunAt === 'never') {
            return 99;
        } else if (dayNames.includes(aiRunAt)) {
            return dayNames.indexOf(aiRunAt);
        } else {
            let dayOfWeek = new Date(aiRunAt).getDay();
            return dayOfWeek;
        }
    } else {
        return -1;
    }
}

function prepareDateRange(s) {
    let fromDateObj, toDateObj;

    // Check if fromDate and toDate are defined and valid in the sheet
    const dateSwitch = s.fromDate && s.toDate;

    if (dateSwitch) {
        // Use user-provided dates from the sheet. Assumes d/m/yyyy format.
        try {
            const fromParts = s.fromDate.split('/');
            const toParts = s.toDate.split('/');
            fromDateObj = new Date(parseInt(fromParts[2], 10), parseInt(fromParts[1], 10) - 1, parseInt(fromParts[0], 10));
            toDateObj = new Date(parseInt(toParts[2], 10), parseInt(toParts[1], 10) - 1, parseInt(toParts[0], 10));

            // A simple check to see if the dates are valid
            if (isNaN(fromDateObj.getTime()) || isNaN(toDateObj.getTime())) {
                throw new Error("Invalid date format in sheet. Please use d/m/yyyy.");
            }
        } catch (e) {
            Logger.log(`Error parsing dates from sheet: ${e.message}. Falling back to default date range.`);
            // Fallback to default logic if parsing fails
            const today = new Date();
            toDateObj = new Date();
            toDateObj.setDate(today.getDate() - 1); // Yesterday
            fromDateObj = new Date();
            fromDateObj.setDate(today.getDate() - s.numberOfDays);
        }
    } else {
        // Fallback to numberOfDays if no dates are provided
        const today = new Date();
        toDateObj = new Date();
        toDateObj.setDate(today.getDate() - 1); // Yesterday
        fromDateObj = new Date();
        fromDateObj.setDate(today.getDate() - s.numberOfDays);
    }

    const fFromDate = Utilities.formatDate(fromDateObj, s.timezone, 'yyyy-MM-dd');
    const fToDate = Utilities.formatDate(toDateObj, s.timezone, 'yyyy-MM-dd');

    const dateRangeQueryString = `segments.date BETWEEN "${fFromDate}" AND "${fToDate}"`;

    return {
        mainDateRange: dateRangeQueryString,
        dateRangeString: dateRangeQueryString,
        fromDate: fromDateObj,
        toDate: toDateObj
    };
}

function defineElements(s) {
    return {
        impr: ' metrics.impressions ', // metrics
        clicks: ' metrics.clicks ',
        cost: ' metrics.cost_micros ',
        engage: ' metrics.engagements ',
        inter: ' metrics.interactions ',
        conv: ' metrics.conversions ',
        value: ' metrics.conversions_value ',
        allConv: ' metrics.all_conversions ',
        allValue: ' metrics.all_conversions_value ',
        views: ' metrics.video_views ',
        cpv: ' metrics.average_cpv ',
        eventTypes: ' metrics.interaction_event_types ',
        segDate: ' segments.date ', // segments
        prodTitle: ' segments.product_title ',
        prodID: ' segments.product_item_id ',
        aIdCamp: ' segments.asset_interaction_target.asset ',
        interAsset: ' segments.asset_interaction_target.interaction_on_this_asset ',
        campName: ' campaign.name ', // campaign
        campId: ' campaign.id ',
        chType: ' campaign.advertising_channel_type ',
        campUrlOptOut: ' campaign.url_expansion_opt_out ',
        lpResName: ' landing_page_view.resource_name ', // landing page
        lpUnexpUrl: ' landing_page_view.unexpanded_final_url ',
        aIdAsset: ' asset.resource_name ', // asset
        aId: ' asset.id ',
        assetSource: ' asset.source ',
        assetName: ' asset.name ',
        assetText: ' asset.text_asset.text ',
        adUrl: ' asset.image_asset.full_size.url ',
        imgHeight: ' asset.image_asset.full_size.height_pixels ',
        imgWidth: ' asset.image_asset.full_size.width_pixels ',
        imgMime: ' asset.image_asset.mime_type ',
        ytTitle: ' asset.youtube_video_asset.youtube_video_title ',
        ytId: ' asset.youtube_video_asset.youtube_video_id ',
        agId: ' asset_group.id ', // asset group
        assetFtype: ' asset_group_asset.field_type ',
        adPmaxPerf: ' asset_group_asset.performance_label ',
        agStrength: ' asset_group.ad_strength ',
        agStatus: ' asset_group.status ',
        agPrimary: ' asset_group.primary_status ',
        asgName: ' asset_group.name ',
        lgType: ' asset_group_listing_group_filter.type ',
        placement: ' detail_placement_view.group_placement_target_url ', // placement
        placeType: ' detail_placement_view.placement_type ',
        chgDateTime: ' change_event.change_date_time ', // change event
        chgResType: ' change_event.change_resource_type ',
        chgFields: ' change_event.changed_fields ',
        clientTyp: ' change_event.client_type ',
        feed: ' change_event.feed ',
        feedItm: ' change_event.feed_item ',
        newRes: ' change_event.new_resource ',
        oldRes: ' change_event.old_resource ',
        resChgOp: ' change_event.resource_change_operation ',
        resName: ' change_event.resource_name ',
        userEmail: ' change_event.user_email ',

        // New elements for geo performance
        geoViewName: ' geographic_view.resource_name ',
        geoLocationType: ' geographic_view.location_type ',
        geoCountryId: ' geographic_view.country_criterion_id ',
        geoTargetCity: ' segments.geo_target_city ',
        geoTargetRegion: ' segments.geo_target_region ',
        geoTargetState: ' segments.geo_target_state ',
        geoTargetCanton: ' segments.geo_target_canton ',
        geoTargetProvince: ' segments.geo_target_province ',
        geoTargetCounty: ' segments.geo_target_county ',
        geoTargetDistrict: ' segments.geo_target_district ',
        geoTargetTypePos: ' campaign.geo_target_type_setting.positive_geo_target_type ',
        geoTargetTypeNeg: ' campaign.geo_target_type_setting.negative_geo_target_type ',

        // New elements for placements
        pmaxPlacementName: ' performance_max_placement_view.display_name ',
        pmaxPlacement: ' performance_max_placement_view.placement ',
        pmaxPlacementType: ' performance_max_placement_view.placement_type ',
        pmaxPlacementResource: ' performance_max_placement_view.resource_name ',
        pmaxPlacementUrl: ' performance_max_placement_view.target_url ',

        // Demand Gen specific elements
        dgAdId: ' ad_group_ad.ad.id ',
        dgAdName: ' ad_group_ad.ad.name ',
        dgAdType: ' ad_group_ad.ad.type ',
        dgAdStatus: ' ad_group_ad.status ',
        dgAdGroupName: ' ad_group.name ',
        dgAdGroupId: ' ad_group.id ',
        dgFinalUrls: ' ad_group_ad.ad.final_urls ',
        dgAssetId: ' asset.id ',
        dgAssetName: ' asset.name ',
        dgAssetType: ' asset.type ',
        dgImageFileSize: ' asset.image_asset.file_size ',
        dgImageMimeType: ' asset.image_asset.mime_type ',
        dgImageUrl: ' asset.image_asset.full_size.url ',
        dgImageWidth: ' asset.image_asset.full_size.width_pixels ',
        dgImageHeight: ' asset.image_asset.full_size.height_pixels ',
        dgTextContent: ' asset.text_asset.text ',
        dgYouTubeId: ' asset.youtube_video_asset.youtube_video_id ',
        dgYouTubeTitle: ' asset.youtube_video_asset.youtube_video_title ',
        dgCTA: ' asset.call_to_action_asset.call_to_action ',

        // Platform comparable metrics
        platConv: ' metrics.platform_comparable_conversions ',
        platConvByDate: ' metrics.platform_comparable_conversions_by_conversion_date ',
        platConvIntRate: ' metrics.platform_comparable_conversions_from_interactions_rate ',
        platConvIntValPerInt: ' metrics.platform_comparable_conversions_from_interactions_value_per_interaction ',
        platConvVal: ' metrics.platform_comparable_conversions_value ',
        platConvValByDate: ' metrics.platform_comparable_conversions_value_by_conversion_date ',
        platConvValPerCost: ' metrics.platform_comparable_conversions_value_per_cost ',

        // Shopping product elements
        shoppingProductId: ' shopping_product.item_id ',
        shoppingProductTitle: ' shopping_product.title ',
        shoppingProductBrand: ' shopping_product.brand ',
        shoppingProductPrice: ' shopping_product.price_micros ',
        shoppingProductCondition: ' shopping_product.condition ',
        shoppingProductAvailability: ' shopping_product.availability ',

        // filters
        networkType: ' segments.ad_network_type ',
        pMaxOnly: ' AND campaign.advertising_channel_type = "PERFORMANCE_MAX" ', // WHERE
        pMaxShop: ' AND campaign.advertising_channel_type IN ("SHOPPING","PERFORMANCE_MAX") ',
        dgOnly: ' AND campaign.advertising_channel_type = "DEMAND_GEN" ',
        dgShop: ' AND campaign.advertising_channel_type IN ("DEMAND_GEN","SHOPPING","PERFORMANCE_MAX") ',
        campLike: ` AND campaign.name LIKE "%${s.campFilter}%" `,
        agFilter: ' AND asset_group_listing_group_filter.type != "SUBDIVISION" ',
        notInter: ' AND segments.asset_interaction_target.interaction_on_this_asset != "TRUE" ',
        impr0: ' AND metrics.impressions > 0 ',
        cost0: ' AND metrics.cost_micros > 0 ',
        order: ' ORDER BY campaign.name ',
    };
}

function buildQueries(e, s, date) {
    let queries = {};

    queries.assetGroupAssetQuery = 'SELECT ' + [e.campName, e.asgName, e.agId, e.aIdAsset, e.assetFtype, e.campId, e.adPmaxPerf].join(',') +
        ' FROM asset_group_asset ' +
        ' WHERE asset_group_asset.field_type NOT IN ( "BUSINESS_NAME", "CALL_TO_ACTION_SELECTION")' + e.pMaxShop + e.campLike; // remove "HEADLINE", "DESCRIPTION", "LONG_HEADLINE", "LOGO", "LANDSCAPE_LOGO",

    queries.displayVideoQuery = 'SELECT ' + [e.segDate, e.campName, e.aIdCamp, e.cost, e.conv, e.value, e.views, e.cpv, e.impr, e.clicks, e.chType, e.interAsset, e.campId].join(',') +
        ' FROM campaign  WHERE ' + date + e.pMaxShop + e.campLike + e.notInter + e.order;

    queries.assetGroupQuery = 'SELECT ' + [e.segDate, e.campName, e.asgName, e.agStrength, e.agStatus, e.lgType, e.impr, e.clicks, e.cost, e.conv, e.value].join(',') +
        ' FROM asset_group_product_group_view WHERE ' + date + e.agFilter + e.campLike;

    queries.campaignQuery = 'SELECT ' + [e.segDate, e.campName, e.cost, e.conv, e.value, e.views, e.cpv, e.impr, e.clicks, e.chType, e.campId].join(',') +
        ' FROM campaign WHERE ' + date + e.pMaxShop + e.campLike + e.order;

    queries.assetQuery = 'SELECT ' + [e.aIdAsset, e.assetSource, e.ytTitle, e.ytId, e.assetName, e.adUrl, e.assetText, e.imgHeight, e.imgWidth, e.imgMime].join(',') +
        ' FROM asset '

    queries.assetGroupMetricsQuery = 'SELECT ' + [e.campName, e.asgName, e.cost, e.conv, e.value, e.impr, e.clicks, e.agPrimary, e.agStatus, e.agStrength].join(',') +
        ' FROM asset_group WHERE ' + date + e.impr0;

    queries.changeQuery = 'SELECT ' + [e.campName, e.userEmail, e.chgDateTime, e.chgResType, e.chgFields, e.clientTyp, e.feed, e.feedItm, e.newRes, e.oldRes, e.resChgOp].join(',') +
        ' FROM change_event ' +
        ' WHERE change_event.change_date_time DURING LAST_14_DAYS ' + e.pMaxShop + e.campLike +
        ' ORDER BY change_event.change_date_time DESC ' +
        ' LIMIT 9999';

    queries.lpQuery = 'SELECT ' + [e.lpUnexpUrl, e.impr, e.clicks, e.cost, e.conv, e.value, e.chType].join(',') +
        ' FROM landing_page_view WHERE ' + date + e.pMaxOnly + ' ORDER BY metrics.impressions DESC';

    queries.placeQuery = 'SELECT ' + [e.campName, e.placement, e.placeType, e.impr, e.inter, e.views, e.cost, e.conv, e.value].join(',') +
        ' FROM detail_placement_view WHERE ' + date + e.campLike + ' ORDER BY metrics.impressions DESC ';

    queries.productQuery = 'SELECT ' + [e.prodTitle, e.prodID, e.cost, e.conv, e.value, e.impr, e.clicks, e.campName, e.chType].join(',') +
        ' FROM shopping_performance_view WHERE metrics.impressions > 0 AND ' + date + e.pMaxShop + e.campLike;
    // Modify the productQuery based on the value of lotsProducts
    switch (s.lotsProducts) {
        case 1:
            queries.productQuery += ' AND metrics.cost_micros > 0';
            break;
        case 2:
            queries.productQuery += ' AND metrics.conversions > 0';
            break;
    }

    queries.geoPerformanceQuery = 'SELECT ' + [e.geoLocationType, e.geoCountryId, e.geoTargetState, e.geoTargetRegion, e.geoTargetCity,
    e.campName, e.campId, e.impr, e.clicks, e.cost, e.conv, e.value, e.chType].join(',') +
        ' FROM geographic_view WHERE ' + date + e.pMaxOnly + e.cost0;

    queries.locationDataQuery = 'SELECT ' + [e.campName, e.campId, e.geoTargetTypePos, e.geoTargetTypeNeg, e.impr, e.clicks, e.cost, e.conv, e.value].join(',') +
        ' FROM campaign WHERE ' + date + e.pMaxOnly + e.impr0;

    queries.pmaxPlacementQuery = 'SELECT ' + [e.pmaxPlacementName, e.pmaxPlacement, e.pmaxPlacementType, e.pmaxPlacementResource, e.pmaxPlacementUrl, e.campName, e.impr].join(',') +
        ' FROM performance_max_placement_view WHERE ' + date + e.impr0;


    queries.shoppingProductsQuery = 'SELECT ' + [e.shoppingProductId, e.shoppingProductTitle, e.shoppingProductBrand, e.shoppingProductPrice, e.shoppingProductCondition, e.shoppingProductAvailability, e.impr, e.clicks, e.cost, e.conv, e.value].join(',') +
        ' FROM shopping_product WHERE ' + date + e.impr0 + ' AND metrics.clicks > 0';

    return queries;
}
//#endregion

//#region prep output
function getData(q, s) {
    let campaignData = fetchData(q.campaignQuery);
    if (campaignData.length === 0) {
        Logger.log('No eligible PMax campaigns found. Exiting script.');
        return null; // Return null to indicate no eligible campaigns were found
    }
    findTopCampaign(campaignData, s);

    Logger.log('Fetching campaign data. This may take a few moments.');

    let assetGroupAssetData = fetchData(q.assetGroupAssetQuery);
    let displayVideoData = fetchData(q.displayVideoQuery);
    let assetGroupData = fetchData(q.assetGroupQuery);
    let assetData = fetchData(q.assetQuery);
    let assetGroupNewData = fetchData(q.assetGroupMetricsQuery);
    let idAccData = fetchProductData(q.productQuery, s.tCost, s.tRoas, 'idAccount');
    let titleAccountData = fetchProductData(q.productQuery, s.tCost, s.tRoas, 'pTitle');
    let idCountData = fetchProductData(q.productQuery, s.tCost, s.tRoas, 'idCount');
    let placementData = fetchData(q.placeQuery);
    let shoppingProductsData = fetchData(q.shoppingProductsQuery);

    return {
        campaignData: campaignData || [],
        assetGroupAssetData: assetGroupAssetData || [],
        displayVideoData: displayVideoData || [],
        assetGroupData: assetGroupData || [],
        assetData: assetData || [],
        assetGroupNewData: assetGroupNewData || [],
        idAccData: idAccData || [],
        titleAccountData: titleAccountData || [],
        idCountData: idCountData || [],
        placementData: placementData || [],
        shoppingProductsData: shoppingProductsData || []
    };
}

function findTopCampaign(campaignData, s) {
    // find the campaign with the highest cost
    let topCampaign = campaignData.reduce((maxCostCampaign, currentCampaign) => {
        let currentCost = parseInt(currentCampaign['metrics.costMicros'], 10);
        let maxCost = parseInt(maxCostCampaign['metrics.costMicros'], 10);
        return currentCost > maxCost ? currentCampaign : maxCostCampaign;
    });
    s.topCampaign = topCampaign['campaign.name'];
}

function processAndAggData(data, ss, s, mainDateRange) {
    Logger.log('Starting data processing.');

    // Extract marketing assets & de-dupe
    let { displayAssetData, videoAssetData, headlineAssetData, descriptionAssetData, longHeadlineAssetData, logoAssetData } = extractAndFilterData(data);

    // Process and aggregate data
    const processAndAggregate = (dataset, type) => {
        let aggregated = aggregateDataByDateAndCampaign(dataset);
        let metrics = aggregateMetricsByAsset(dataset);
        let enriched = enrichAssetMetrics(metrics, data.assetData, type);
        return { aggregated, metrics, enriched };
    };

    let displayData = processAndAggregate(displayAssetData, 'display');
    let videoData = processAndAggregate(videoAssetData, 'video');
    let headlineData = processAndAggregate(headlineAssetData, 'headline');
    let descriptionData = processAndAggregate(descriptionAssetData, 'description');
    let longHeadlineData = processAndAggregate(longHeadlineAssetData, 'long_headline');
    let logoData = processAndAggregate(logoAssetData, 'logo');

    // Process campaign and asset group data
    let processedCampData = processData(data.campaignData);
    let processedAssetGroupData = processData(data.assetGroupData);
    let dAGData = tidyAssetGroupData(data.assetGroupNewData);

    // Combine all non-search metrics, calc 'search' & process summary
    let nonSearchData = [...displayData.aggregated, ...videoData.aggregated, ...processedAssetGroupData];
    let searchResults = getSearchResults(processedCampData, nonSearchData);
    let dTotal = processTotalData(processedCampData, processedAssetGroupData, displayData.aggregated, videoData.aggregated, searchResults);
    let dSummary = processSummaryData(processedCampData, processedAssetGroupData, displayData.aggregated, videoData.aggregated, searchResults);

    // Merge assets with details
    let dAssets = mergeAssetsWithDetails(
        displayData.metrics, displayData.enriched,
        videoData.metrics, videoData.enriched,
        headlineData.metrics, headlineData.enriched,
        descriptionData.metrics, descriptionData.enriched,
        longHeadlineData.metrics, longHeadlineData.enriched,
        logoData.metrics, logoData.enriched
    );

    // Extract terms and n-grams
    let terms = extractSearchCats(ss, mainDateRange, s);
    let totalTerms = aggregateTerms(terms, ss);
    let sNgrams = extractSearchNgrams(s, terms, 'search');
    let tNgrams = extractTitleNgrams(s, data.titleAccountData, 'title');
    let placements = processPlacementData(data.placementData, s);
    let idCount = processIdCountData(data.idCountData);
    let dShoppingProducts = processShoppingProductsData(data.shoppingProductsData);

    // Return all data
    return {
        dAssets,
        dAGData,
        dTotal,
        dSummary,
        terms,
        totalTerms,
        sNgrams,
        tNgrams,
        placements,
        idCount,
        dShoppingProducts
    }
}

function writeAllDataToSheets(ss, data) {
    Logger.log('Writing data to sheets.');

    let writeOperations = [

        { sheetName: 'asset', data: data.dAssets, outputType: 'clear' },
        { sheetName: 'totals', data: data.dTotal, outputType: 'clear' },
        { sheetName: 'summary', data: data.dSummary, outputType: 'clear' },
        { sheetName: 'groups', data: data.dAGData, outputType: 'clear' },
        { sheetName: 'pTitle', data: data.titleAccountData, outputType: 'clear' },
        { sheetName: 'ID', data: data.idAccData, outputType: 'clear' },
        { sheetName: 'terms', data: data.terms, outputType: 'clear' },
        { sheetName: 'totalTerms', data: data.totalTerms, outputType: 'clear' },
        { sheetName: 'tNgrams', data: data.tNgrams, outputType: 'clear' },
        { sheetName: 'sNgrams', data: data.sNgrams, outputType: 'clear' },
        { sheetName: 'placement', data: data.placements, outputType: 'clear' },
        { sheetName: 'idCount', data: data.idCount, outputType: 'clear' },
        // { sheetName: 'shoppingProducts', data: data.dShoppingProducts, outputType: 'clear' },

    ];

    let batchOperations = [];

    writeOperations.forEach(operation => {
        let { sheetName, data, outputType } = operation;
        let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

        if (sheet.getMaxRows() > 1) {
            if (outputType === 'notLast') {
                sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent();
            } else if (outputType === 'clear') {
                sheet.clearContents();
            }
        }

        if (data && data.length > 0) {
            let outputData = prepareOutputData(data, outputType);
            let numRows = outputData.length;
            let numColumns = outputData[0].length;
            batchOperations.push({ sheet, numRows, numColumns, outputData });
        }
    });

    batchOperations.forEach(({ sheet, numRows, numColumns, outputData }) => {
        try {
            let range = sheet.getRange(1, 1, numRows, numColumns);
            range.setValues(outputData);
        } catch (e) {
            console.error(`Error writing to ${sheet.getName()}: ${e.message}`);
            console.error(`Data structure: ${JSON.stringify(outputData.slice(0, 2))}`);
        }
    });

    SpreadsheetApp.flush();
    Logger.log('Data writing to sheets completed.');
}

function prepareOutputData(data) {
    if (!Array.isArray(data)) {
        Logger.log('Warning: Data is expected to be an array, but received an object.');
        return []; // Return an empty array to avoid issues
    }

    if (data.length === 0) {
        return []; // Return an empty array if data is empty
    }

    if (Array.isArray(data[0])) {
        // Data is already an array of arrays
        return data;
    } else {
        // Data is an array of objects
        let headers = Object.keys(data[0]);
        let values = data.map(row => headers.map(header => row[header] !== null && row[header] !== undefined ? row[header] : ""));
        return [headers].concat(values);
    }
}

function outputDataToSheet(ss, sheetName, data, outputType) {
    let startTime = new Date();

    // Create or access the sheet
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    // Check if data is undefined or empty, and if so, log the issue and return
    if (!data || data.length === 0) {
        Logger.log('Maybe try a different date range. Currently, there is no data to write to ' + sheetName);
        if (outputType === 'notLast') {
            sheet.getRange(2, 1, sheet.getMaxRows(), sheet.getMaxColumns() - 1).clearContent();
        } else if (outputType === 'clear') {
            sheet.clearContents();
        }
        return;
    }

    // Determine the number of columns in the data
    let numberOfColumns = Array.isArray(data[0]) ? data[0].length : Object.keys(data[0]).length;

    // Clear the dynamic range based on the data length
    sheet.getRange(1, 1, sheet.getMaxRows(), numberOfColumns).clearContent();

    // Prepare the output data
    let outputData;
    if (!Array.isArray(data[0])) {
        let headers = Object.keys(data[0]);
        let values = data.map(row => headers.map(header => row[header] ?? ""));
        outputData = [headers].concat(values);
    } else {
        outputData = data;
    }

    // Write all data to the sheet
    sheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

}

function extractSearchCats(ss, mainDateRange, s) {
    let campaignIdsQuery = AdsApp.report(`
    SELECT campaign.id, campaign.name, metrics.clicks, 
    metrics.impressions, metrics.conversions, metrics.conversions_value
    FROM campaign 
    WHERE campaign.status != 'REMOVED' 
    AND campaign.advertising_channel_type = "PERFORMANCE_MAX" 
    AND metrics.impressions > 0 AND ${mainDateRange} 
    ORDER BY metrics.conversions DESC `
    );

    let rows = campaignIdsQuery.rows();
    let allSearchTerms = [['Campaign Name', 'Campaign ID', 'Category Label', 'Clicks', 'Impr', 'Conv', 'Value', 'Bucket', 'Distance']];

    while (rows.hasNext()) {
        let row = rows.next();
        let campaignName = row['campaign.name'];
        let campaignId = row['campaign.id'];

        let query = AdsApp.report(` 
        SELECT campaign_search_term_insight.category_label, metrics.clicks, 
        metrics.impressions, metrics.conversions, metrics.conversions_value  
        FROM campaign_search_term_insight 
        WHERE ${mainDateRange}
        AND campaign_search_term_insight.campaign_id = ${campaignId} 
        ORDER BY metrics.impressions DESC `
        );

        let searchTermRows = query.rows();
        while (searchTermRows.hasNext()) {
            let row = searchTermRows.next();
            let term = (row['campaign_search_term_insight.category_label'] || 'blank').toLowerCase();
            let { bucket, distance } = determineBucketAndDistance(term, s.brandTerm, s.levenshtein);
            term = cleanNGram(term);
            allSearchTerms.push([campaignName, campaignId,
                term,
                row['metrics.clicks'],
                row['metrics.impressions'],
                row['metrics.conversions'],
                row['metrics.conversions_value'],
                bucket,
                distance]);
        }
    }

    return allSearchTerms;
}

function aggregateTerms(searchTerms, ss) {
    let aggregated = {}; // { term: { clicks: 0, impr: 0, conv: 0, value: 0, bucket: '', distance: 0 }, ... }

    for (let i = 1; i < searchTerms.length; i++) { // Start from 1 to skip headers
        let term = searchTerms[i][2] || 'blank'; // Use 'blank' for empty search terms

        if (!aggregated[term]) {
            aggregated[term] = {
                clicks: 0,
                impr: 0,
                conv: 0,
                value: 0,
                bucket: searchTerms[i][7],  // Assuming bucket is in the 8th position of your array
                distance: searchTerms[i][8]  // Assuming distance is in the 9th position of your array
            };
        }

        aggregated[term].clicks += Number(searchTerms[i][3]);
        aggregated[term].impr += Number(searchTerms[i][4]);
        aggregated[term].conv += Number(searchTerms[i][5]);
        aggregated[term].value += Number(searchTerms[i][6]);
        // Assuming that the bucket and distance are the same for all instances of a term,
        // we don't aggregate them but just take the value from the first instance.
    }

    let totalTerms = [['Category Label', 'Clicks', 'Impr', 'Conv', 'Value', 'Bucket', 'Distance']]; // Header row
    for (let term in aggregated) {
        totalTerms.push([
            term,
            aggregated[term].clicks,
            aggregated[term].impr,
            aggregated[term].conv,
            aggregated[term].value,
            aggregated[term].bucket,  // Adding bucket to output
            aggregated[term].distance  // Adding distance to output
        ]);
    }

    let header = totalTerms.shift(); // Remove the header before sorting
    // Sort by impressions descending
    totalTerms.sort((a, b) => b[2] - a[2]);
    totalTerms.unshift(header); // Prepend the header back to the top

    return totalTerms;

}

function extractSearchNgrams(s, data, type) {
    let nGrams = {};

    data.slice(1).forEach(row => {
        let term = type === 'search' ? row[2] : (row['Product Title'] ? row['Product Title'].toLowerCase() : '');
        let terms = term.split(' ');
        terms.forEach((nGram) => {
            nGram = nGram || 'blank';
            nGram = cleanNGram(nGram);
            if (!nGrams[nGram]) {
                nGrams[nGram] = {
                    nGram: nGram,
                    impr: 0,
                    clicks: 0,
                    conv: 0,
                    value: 0,
                    cost: type === 'title' ? 0 : undefined
                };
            }
            nGrams[nGram].impr += type === 'search' ? Number(row[4]) : row['Impr'];
            nGrams[nGram].clicks += type === 'search' ? Number(row[3]) : row['Clicks'];
            nGrams[nGram].conv += type === 'search' ? Number(row[5]) : row['Conv'];
            nGrams[nGram].value += type === 'search' ? Number(row[6]) : row['Value'];
            if (type === 'title') {
                nGrams[nGram].cost += row['Cost'];
            }
        });
    });

    let allNGrams = type === 'search'
        ? [['nGram', 'Impr', 'Clicks', 'Conv', 'Value', 'CTR', 'CvR', 'AOV', 'Bucket']]
        : [['nGram', 'Impr', 'Clicks', 'Cost', 'Conv', 'Value', 'CTR', 'CvR', 'AOV', 'ROAS', 'Bucket']];

    let nGramsList = Object.values(nGrams);

    let totalClicks = nGramsList.reduce((sum, i) => sum + i.clicks, 0);
    let totalConversions = nGramsList.reduce((sum, i) => sum + i.conv, 0);

    let percentile80Clicks = 0.8 * totalClicks;
    let percentile80Conversions = 0.8 * totalConversions;

    let cumulativeClicks = 0;
    let cumulativeConversions = 0;
    let clicks80Percentile = 0;
    let conv80Percentile = 0;

    nGramsList.sort((a, b) => b.clicks - a.clicks);
    for (let i of nGramsList) {
        cumulativeClicks += i.clicks;
        if (cumulativeClicks <= percentile80Clicks) {
            clicks80Percentile = i.clicks;
        } else {
            break;
        }
    }

    nGramsList.sort((a, b) => b.conv - a.conv);
    for (let i of nGramsList) {
        cumulativeConversions += i.conv;
        if (cumulativeConversions <= percentile80Conversions) {
            conv80Percentile = i.conv;
        } else {
            break;
        }
    }

    for (let term in nGrams) {
        let item = nGrams[term];
        item.CTR = item.impr > 0 ? item.clicks / item.impr : 0;
        item.CvR = item.clicks > 0 ? item.conv / item.clicks : 0;
        item.AOV = item.conv > 0 ? item.value / item.conv : 0;

        if (item.clicks === 0) {
            item.bucket = 'zombie';
        } else if (item.conv === 0) {
            item.bucket = 'zeroconv';
        } else if (item.clicks < clicks80Percentile) {
            item.bucket = item.conv < conv80Percentile ? 'Lclicks_Lconv' : 'Lclicks_Hconv';
        } else {
            item.bucket = item.conv < conv80Percentile ? 'Hclicks_Lconv' : 'Hclicks_Hconv';
        }

        if (type === 'title') {
            item.ROAS = item.cost > 0 ? item.value / item.cost : 0;
            allNGrams.push([item.nGram, item.impr, item.clicks, item.cost, item.conv, item.value, item.CTR, item.CvR, item.AOV, item.ROAS, item.bucket]);
        } else {
            allNGrams.push([item.nGram, item.impr, item.clicks, item.conv, item.value, item.CTR, item.CvR, item.AOV, item.bucket]);
        }
    }

    allNGrams.sort((a, b) => {
        if (a[0] === 'nGram') return -1;
        if (b[0] === 'nGram') return 1;
        return type === 'search' ? b[1] - a[1] : b[3] - a[3]; // Sort by Clicks for search and Cost for title
    });

    let brandTerm = s.brandTerm.includes(',') ? s.brandTerm.split(/[ ,]+/).map(i => i.toLowerCase()) : s.brandTerm.split(' ').map(i => i.toLowerCase());
    allNGrams = allNGrams.filter(i => !brandTerm.includes(i[0]) && i[0] !== 'blank');

    return allNGrams;
}

function extractTitleNgrams(s, data, type) {
    let nGrams = {};

    data.slice(1).forEach(row => {
        let term = type === 'search' ? row[2] : (row['Product Title'] ? row['Product Title'].toLowerCase() : '');
        let terms = term.split(' ');

        terms.forEach((nGram) => {
            nGram = nGram || 'blank';
            nGram = cleanNGram(nGram);
            if (!nGrams[nGram]) {
                nGrams[nGram] = {
                    nGram: nGram,
                    impr: 0,
                    clicks: 0,
                    conv: 0,
                    value: 0,
                    cost: type === 'title' ? 0 : undefined
                };
            }
            nGrams[nGram].impr += type === 'search' ? Number(row[4]) : row['Impr'];
            nGrams[nGram].clicks += type === 'search' ? Number(row[3]) : row['Clicks'];
            nGrams[nGram].conv += type === 'search' ? Number(row[5]) : row['Conv'];
            nGrams[nGram].value += type === 'search' ? Number(row[6]) : row['Value'];
            if (type === 'title') {
                nGrams[nGram].cost += row['Cost'];
            }
        });
    });

    let allNGrams = type === 'search'
        ? [['nGram', 'Impr', 'Clicks', 'Conv', 'Value', 'CTR', 'CvR', 'AOV', 'Bucket']]
        : [['nGram', 'Impr', 'Clicks', 'Cost', 'Conv', 'Value', 'CTR', 'CvR', 'AOV', 'ROAS', 'Bucket']];

    for (let term in nGrams) {
        let i = nGrams[term];
        i.CTR = i.impr > 0 ? i.clicks / i.impr : 0;
        i.CvR = i.clicks > 0 ? i.conv / i.clicks : 0;
        i.AOV = i.conv > 0 ? i.value / i.conv : 0;
        if (type === 'title') {
            i.ROAS = i.cost > 0 ? i.value / i.cost : 0;
            i.bucket = determineBucket(i.cost, i.conv, i.ROAS, (s.tCost * 10), (s.tRoas * 2));
        }
        if (type === 'search') {
            i.bucket = determineSearchBucket(i.impr, i.clicks, i.conv, i.value);
        }
        allNGrams.push(type === 'search'
            ? [i.nGram, i.impr, i.clicks, i.conv, i.value, i.CTR, i.CvR, i.AOV, i.bucket]
            : [i.nGram, i.impr, i.clicks, i.cost, i.conv, i.value, i.CTR, i.CvR, i.AOV, i.ROAS, i.bucket]
        );
    }

    allNGrams.sort((a, b) => {
        if (a[0] === 'nGram') return -1;
        if (b[0] === 'nGram') return 1;
        return type === 'search' ? b[2] - a[2] : b[3] - a[3]; // Sort by Impr for search and Cost for title
    });

    let brandTerm = s.brandTerm.includes(',') ? s.brandTerm.split(/[ ,]+/).map(i => i.toLowerCase()) : s.brandTerm.split(' ').map(i => i.toLowerCase());
    allNGrams = allNGrams.filter(i => !brandTerm.includes(i[0]) && i[0] !== 'blank' && i[0] !== '');

    return allNGrams;
}

function adjustSheetVisibilityBasedOnAccountType(ss, s) {
    Logger.log('Adjusting visible tabs based on account type.');
    try {
        let leadgenSheets = ['Account.', 'Campaign.', 'Categories.', 'Asset Groups.']; // no alt tab for Assets or Pages yet
        let ecommerceSheets = ['Account', 'Campaign', 'Categories', 'Title', 'nGram', 'Comp', 'ID', 'Asset Groups'];

        // Determine which sheets to show/hide based on account type
        let sheetsToShow = s.accountType === 'leadgen' ? leadgenSheets : ecommerceSheets;
        let sheetsToHide = s.accountType === 'leadgen' ? ecommerceSheets : leadgenSheets;

        // Show and hide sheets accordingly
        sheetsToShow.forEach(sheetName => ss.getSheetByName(sheetName)?.showSheet());
        sheetsToHide.forEach(sheetName => ss.getSheetByName(sheetName)?.hideSheet());

        // Adjusting rows visibility in 'Advanced' tab based on account type
        let advancedSheet = ss.getSheetByName('Advanced');
        if (s.accountType === 'leadgen') {
            advancedSheet.hideRows(12, 6); // Hides rows 12 to 17
            advancedSheet.hideRows(26, 5); // Hides rows 26 to 30
        } else {
            advancedSheet.showRows(12, 6); // Shows rows 12 to 17
            advancedSheet.showRows(26, 5); // Shows rows 26 to 30
        }

        try {
            let selectedCampaignRange = ss.getRangeByName('selectedCampaign');
            if (selectedCampaignRange) {
                selectedCampaignRange.setValue(s.topCampaign);
            }
        } catch (e) {
            Logger.log('selectedCampaign named range not found: ' + e.message);
        }
        
        try {
            let selectedCampaignDotRange = ss.getRangeByName('selectedCampaign.');
            if (selectedCampaignDotRange) {
                selectedCampaignDotRange.setValue(s.topCampaign);
            }
        } catch (e) {
            Logger.log('selectedCampaign. named range not found: ' + e.message);
        }
    } catch (error) {
        console.error('Error adjusting sheet visibility:', error.message);
    }
}

function formatDateLiteral(dateString) {
    try {
        // Split the date string assuming d/m/yyyy format based on the validation in updateVariablesFromSheet
        const parts = dateString.split('/');
        if (parts.length !== 3) {
            throw new Error(`Invalid date parts after split for: ${dateString}`);
        }

        let day = parts[0];
        let month = parts[1];
        const year = parts[2];

        // Ensure year looks like a 4-digit year (basic check)
        if (year.length !== 4 || isNaN(parseInt(year))) {
            throw new Error(`Invalid year format: ${year} in ${dateString}`);
        }

        // Pad day and month with leading zeros if necessary
        if (day.length === 1) {
            day = '0' + day;
        }
        if (month.length === 1) {
            month = '0' + month;
        }

        // Basic validation for month and day numbers
        const monthNum = parseInt(month);
        const dayNum = parseInt(day);
        if (monthNum < 1 || monthNum > 12 || dayNum < 1 || dayNum > 31) {
            throw new Error(`Invalid day (${dayNum}) or month (${monthNum}) in date: ${dateString}`);
        }


        // Assemble in yyyy-MM-dd format
        const formattedDate = `${year}-${month}-${day}`;
        // Optional: Final check if the assembled date is valid according to Utilities.formatDate
        // This helps catch things like 31/02/2024, though the Ads API would likely reject it anyway.
        try {
            Utilities.formatDate(new Date(formattedDate), "UTC", "yyyy-MM-dd");
        } catch (e) {
            throw new Error(`Assembled date ${formattedDate} from ${dateString} is invalid: ${e.message}`);
        }

        return formattedDate;

    } catch (error) {
        Logger.log(`Error in formatDateLiteral for input "${dateString}": ${error.message}`);
        // Fallback or re-throw depending on desired behavior. Returning undefined might cause issues.
        // It might be better to let the script fail here if formatting fails, indicating bad input.
        throw new Error(`Failed to format date: ${dateString}. Please use d/m/yyyy format. Error: ${error.message}`);
        // return undefined; // Or return a default/fallback if preferred
    }
}

function checkVersion(userScriptVersion, userSheetVersion, ss, template) {
    let CONTROL_SHEET = 'https://docs.google.com/spreadsheets/d/16W7nqAGg7LbqykOCdwsGLB93BMUWvU9ZjlVprIPZBtg/' // used to find latest version
    let controlSheet = safeOpenAndShareSpreadsheet(CONTROL_SHEET);
    // get value of named ranges currentScriptVersion & currentSheetVersion
    let latestScriptVersion, latestSheetVersion;
    try {
        latestScriptVersion = controlSheet.getRangeByName('latestScriptVersion').getValue();
        latestSheetVersion = controlSheet.getRangeByName('latestSheetVersion').getValue();
    } catch (e) {
        console.error(`Error fetching current script/sheet version: ${e.message}`);
    }

    // Display messages based on version comparison & write them to userMessage in the main ss
    let userMessage = '';
    let userMessageRange = ss.getRangeByName('userMessage');

    if (userScriptVersion !== latestScriptVersion || userSheetVersion !== latestSheetVersion) {
        if (userScriptVersion !== latestScriptVersion) {
            userMessage = "Time to update your script. You are using: " + userScriptVersion + ", Latest version is: " + latestScriptVersion + " from here: https://mikerhodes.circle.so/c/script/";
        }
        if (userSheetVersion !== latestSheetVersion) {
            userMessage = "Time to update your sheet. You are using: " + userSheetVersion + ", Latest version is: " + latestSheetVersion;
        }
    } else {
        userMessage = "Great work! You're using the latest versions. Script: " + userScriptVersion + ", Sheet: " + userSheetVersion;
    }
    userMessageRange.setValue(userMessage);
    Logger.log(userMessage);
}

function flattenObject(ob) {
    let toReturn = {};
    let stack = [{ obj: ob, prefix: '' }];

    while (stack.length > 0) {
        let { obj, prefix } = stack.pop();
        for (let i in obj) {
            if (obj.hasOwnProperty(i)) {
                let key = prefix ? prefix + '.' + i : i;
                if (typeof obj[i] === 'object' && obj[i] !== null) {
                    stack.push({ obj: obj[i], prefix: key });
                } else {
                    toReturn[key] = obj[i];
                }
            }
        }
    }

    return toReturn;
}

function fetchData(q) {
    try {
        let data = [];
        let iterator = AdsApp.search(q, { 'apiVersion': 'v20' });
        while (iterator.hasNext()) {
            let row = flattenObject(iterator.next());
            data.push(row); // Flatten the row data
        }
        return data;
    } catch (error) {
        Logger.log(`Error fetching data for query: ${q}`);
        Logger.log(`Error message: ${error.message}`);
        return null; // Return null to indicate that the data fetch failed
    }
}
//#endregion

//#region data processing
function processIdCountData(rawData) {
    const headers = ['Product Title', 'ID', 'inBoth', 'countPmax', 'countShop', 'costPmax', 'costShop'];

    // Transform the data into rows
    const processedRows = rawData.map(item => ({
        'Product Title': item['Product Title'] || '',  // Changed to match the format
        'ID': item.ID || '',
        'inBoth': item.inBoth || 0,
        'countPmax': item.countPmax || 0,
        'countShop': item.countShop || 0,
        'costPmax': item.costPmax || 0,
        'costShop': item.costShop || 0
    }));

    // Sort rows - first by inBoth (1 first), then by costPmax
    const sortedRows = processedRows.sort((a, b) => {
        // First sort by inBoth (1 first)
        if (b.inBoth !== a.inBoth) {
            return b.inBoth - a.inBoth;
        }
        // Then sort by costPmax within each inBoth group
        return b.costPmax - a.costPmax;
    });

    // Convert to array format
    const rows = sortedRows.map(item => [
        item['Product Title'],
        item.ID,
        item.inBoth,
        item.countPmax,
        item.countShop,
        item.costPmax.toFixed(2),
        item.costShop.toFixed(2)
    ]);

    return [headers, ...rows];
}

function processGeoPerformanceData(geoData, locationData, ss, availableGeoFields = []) {
    if (!geoData?.length) return {
        locations: [['No data available']],
        campaigns: [['No data available']]
    };

    // Initialize base structures
    const campaignSettings = new Map(locationData.map(row =>
        [row['campaign.name'], row['campaign.geoTargetTypeSetting.positiveGeoTargetType']]
    ));

    const locationMap = new Map(
        ss.getSheetByName('map')?.getRange('A:C')
            .getValues()
            .filter(([id, name, type]) => id && name && type)
            .map(([id, name, type]) => [`geoTargetConstants/${id.toString()}`, { name, type }]) || []
    );

    // Helper function to create empty stat object
    const createEmptyStat = (type = null) => ({
        costP: 0, costI: 0, convP: 0, convI: 0, valueP: 0, valueI: 0,
        ...(type && { type })
    });

    // Stats storage - create dynamic stats based on available geo fields
    const stats = {
        country: new Map(),
        campaign: new Map()
    };

    // Add stats maps for each available geo field
    availableGeoFields.forEach(field => {
        stats[field.name.toLowerCase()] = new Map();
    });

    // Campaign location tracking
    const campaignLocationCosts = new Map();

    // Helper function to update stats
    const updateStats = (map, key, metrics, isPresence, type = null) => {
        if (!map.has(key)) {
            map.set(key, createEmptyStat(type));
        }
        const stat = map.get(key);
        const suffix = isPresence ? 'P' : 'I';
        Object.entries(metrics).forEach(([metric, value]) => {
            stat[metric + suffix] += value;
        });
    };

    // Process each row of data
    geoData.forEach(row => {
        const isPresence = row['geographicView.locationType'] === 'LOCATION_OF_PRESENCE';
        const metrics = {
            cost: (parseInt(row['metrics.costMicros']) || 0) / 1e6,
            conv: parseFloat(row['metrics.conversions']) || 0,
            value: parseFloat(row['metrics.conversionsValue']) || 0
        };

        const campaignName = row['campaign.name'];

        // Process geographical data dynamically based on available fields
        const geoTypes = [
            {
                type: 'country',
                id: row['geographicView.countryCriterionId'],
                prefix: 'geoTargetConstants/',
                validType: 'Country'
            }
        ];

        // Add dynamic geo types based on available fields
        availableGeoFields.forEach(field => {
            const fieldMapping = {
                'Canton': { segment: 'segments.geoTargetCanton', validTypes: ['Canton'] },
                'Region': { segment: 'segments.geoTargetRegion', validTypes: ['Region'] },
                'State': { segment: 'segments.geoTargetState', validTypes: ['State'] },
                'Province': { segment: 'segments.geoTargetProvince', validTypes: ['Province'] },
                'County': { segment: 'segments.geoTargetCounty', validTypes: ['County'] },
                'District': { segment: 'segments.geoTargetDistrict', validTypes: ['District'] }
            };

            const mapping = fieldMapping[field.name];
            if (mapping && row[mapping.segment]) {
                geoTypes.push({
                    type: field.name.toLowerCase(),
                    id: row[mapping.segment],
                    prefix: '',
                    validType: mapping.validTypes
                });
            }
        });

        geoTypes.forEach(({ type, id, prefix, validType }) => {
            if (id) {
                const lookupKey = `${prefix}${id}`;
                const mapped = locationMap.get(lookupKey);
                const isValidType = Array.isArray(validType)
                    ? validType.includes(mapped?.type)
                    : mapped?.type === validType;

                if (mapped && isValidType) {
                    updateStats(stats[type], mapped.name, metrics, isPresence, mapped.type);

                    // Track country costs for campaigns
                    if (type === 'country') {
                        if (!campaignLocationCosts.has(campaignName)) {
                            campaignLocationCosts.set(campaignName, new Map());
                        }
                        const locationCosts = campaignLocationCosts.get(campaignName);
                        locationCosts.set(mapped.name,
                            (locationCosts.get(mapped.name) || 0) + metrics.cost);
                    }
                }
            }
        });

        // Process campaign stats
        updateStats(stats.campaign, campaignName, metrics, isPresence);
    });

    // Processing helper functions
    const getTopLocation = (campaignName) => {
        const locationCosts = campaignLocationCosts.get(campaignName);
        if (!locationCosts) return { location: 'Unknown', percentage: '0' };

        const totalCost = Array.from(locationCosts.values()).reduce((a, b) => a + b, 0);
        const [topLocation, topCost] = Array.from(locationCosts.entries())
            .reduce((max, curr) => curr[1] > max[1] ? curr : max, ['', 0]);

        return {
            location: topLocation,
            percentage: totalCost > 0 ? (topCost / totalCost * 100).toFixed(0) : '0'
        };
    };

    const formatStats = (stat) => [
        ...['costP', 'costI', 'convP', 'convI', 'valueP', 'valueI'].map(h => stat[h].toFixed(2)),
        (stat.costP > 0 ? stat.valueP / stat.costP : 0).toFixed(2),
        (stat.costI > 0 ? stat.valueI / stat.costI : 0).toFixed(2)
    ];

    // Create output rows with headers
    const headers = ['costP', 'costI', 'convP', 'convI', 'valueP', 'valueI', 'roasP', 'roasI'];

    const createRows = (map, includeSettings = false) => {
        const rows = Array.from(map, ([key, data]) => {
            const row = [key, ...formatStats(data)];

            if (includeSettings) {
                const setting = campaignSettings.get(key);
                row.push(setting === 'PRESENCE' ? 'ok' :
                    setting === 'PRESENCE_OR_INTEREST' ? 'review' : 'unknown');
                const { location, percentage } = getTopLocation(key);
                row.push(location, `${percentage}%`);
            }

            if (data.type) row.push(data.type);
            return row;
        }).sort((a, b) => (parseFloat(b[1]) + parseFloat(b[2])) - (parseFloat(a[1]) + parseFloat(a[2])));

        const nameHeader = includeSettings ? 'Campaign' : 'Location';
        const extraHeaders = includeSettings ?
            ['Setting', 'Top Location', 'Top Location %'] : ['Type'];
        return [[nameHeader, ...headers, ...extraHeaders], ...rows];
    };

    // Create dynamic locations array with all available geo fields
    const locationRows = [['Location', ...headers, 'Type']]; // header row
    locationRows.push(...createRows(stats.country).slice(1)); // skip header

    // Add rows for each available geo field
    availableGeoFields.forEach(field => {
        const fieldType = field.name.toLowerCase();
        if (stats[fieldType]) {
            locationRows.push(...createRows(stats[fieldType]).slice(1)); // skip header
        }
    });

    return {
        locations: locationRows,
        campaigns: createRows(stats.campaign, true)
    };
}

function buildTestQuery(e, geoField, dateRange) {
    const fields = [
        e.geoCountryId,
        geoField,
        e.impr,
        e.chType
    ].join(',');

    return `SELECT ${fields} FROM geographic_view WHERE ${dateRange}${e.pMaxOnly}${e.cost0} LIMIT 1`;
}

function buildTestQueryForCountry(e, geoField, countryId, dateRange) {
    const fields = [
        e.geoCountryId,
        geoField,
        e.impr,
        e.chType
    ].join(',');

    return `SELECT ${fields} FROM geographic_view WHERE ${dateRange} AND ${e.geoCountryId} = ${countryId}${e.pMaxOnly}${e.cost0} LIMIT 1`;
}

function discoverAvailableGeoFields(e, dateRange) {
    // First, discover which countries have data
    const countryQuery = `SELECT ${e.geoCountryId}, ${e.impr}, ${e.chType} FROM geographic_view WHERE ${dateRange}${e.pMaxOnly}${e.cost0}`;
    const countryData = fetchData(countryQuery);

    if (!countryData || countryData.length === 0) {
        return [];
    }

    // Deduplicate countries manually since GROUP BY isn't supported
    const countrySet = new Set();
    countryData.forEach(row => {
        if (row['geographicView.countryCriterionId']) {
            countrySet.add(row['geographicView.countryCriterionId']);
        }
    });
    const countries = Array.from(countrySet);

    const geoFieldsToTest = [
        { name: 'Canton', field: e.geoTargetCanton },
        { name: 'Region', field: e.geoTargetRegion },
        { name: 'State', field: e.geoTargetState },
        { name: 'Province', field: e.geoTargetProvince },
        { name: 'County', field: e.geoTargetCounty },
        { name: 'District', field: e.geoTargetDistrict }
    ];

    // Track which fields are available for which countries
    const fieldCountryMap = new Map();

    // Test each geo field for each country
    countries.forEach(countryId => {

        geoFieldsToTest.forEach(geoField => {
            const testQuery = buildTestQueryForCountry(e, geoField.field, countryId, dateRange);
            const testData = fetchData(testQuery);

            if (testData && testData.length > 0) {

                if (!fieldCountryMap.has(geoField.name)) {
                    fieldCountryMap.set(geoField.name, { field: geoField, countries: new Set() });
                }
                fieldCountryMap.get(geoField.name).countries.add(countryId);
            }
        });
    });

    // Only include fields that are available for ALL countries
    const universalFields = [];
    fieldCountryMap.forEach((fieldInfo, fieldName) => {
        if (fieldInfo.countries.size === countries.length) {
            universalFields.push(fieldInfo.field);
        }
    });

    return universalFields;
}

function buildGeoPerformanceQuery(e, geoFields, dateRange) {
    const fields = [
        e.campName,
        ...geoFields,
        e.geoLocationType,
        e.impr, e.clicks, e.cost, e.conv, e.value, e.chType
    ].join(',');

    const query = `SELECT ${fields} FROM geographic_view WHERE ${dateRange}${e.pMaxOnly}${e.cost0}`;
    return query;
}

function processPlacementData(data, s) {
    if (!data || data.length === 0) {
        Logger.log('No placement data available.');
        return [['No data available']];
    }

    let minPlaceImpr = s.minPlaceImpr;

    // Aggregate ALL data first for "All Campaigns" view
    const aggregatedData = new Map();
    data.forEach(row => {
        const key = row['detailPlacementView.groupPlacementTargetUrl'];
        if (!aggregatedData.has(key)) {
            aggregatedData.set(key, {
                placement: key,
                type: row['detailPlacementView.placementType'],
                impressions: 0, interactions: 0, views: 0, cost: 0, conversions: 0, value: 0
            });
        }
        const entry = aggregatedData.get(key);
        entry.impressions += Number(row['metrics.impressions']) || 0;
        entry.interactions += Number(row['metrics.interactions']) || 0;
        entry.views += Number(row['metrics.videoViews']) || 0;
        entry.cost += (Number(row['metrics.costMicros']) || 0) / 1e6;
        entry.conversions += Number(row['metrics.conversions']) || 0;
        entry.value += Number(row['metrics.conversionsValue']) || 0;
    });

    let processedData = [];
    const headers = [['Campaign', 'Placement', 'Type', 'Impr.', 'Interactions', 'Views', 'Cost', 'Conv.', 'Value', 'CTR', 'CVR', 'CPA', 'ROAS']];

    // Add aggregated "All Campaigns" rows that meet the threshold
    for (const placementData of aggregatedData.values()) {
        if (placementData.impressions >= minPlaceImpr) {
            const cost = placementData.cost;
            const interactions = placementData.interactions;
            const impressions = placementData.impressions;
            const conversions = placementData.conversions;
            const value = placementData.value;
            processedData.push([
                'All Campaigns',
                placementData.placement,
                placementData.type,
                impressions,
                interactions,
                placementData.views,
                cost,
                conversions,
                value,
                impressions > 0 ? interactions / impressions : 0,
                interactions > 0 ? conversions / interactions : 0,
                conversions > 0 ? cost / conversions : 0,
                cost > 0 ? value / cost : 0
            ]);
        }
    }

    // Sort "All campaigns" section
    processedData.sort((a, b) => b[3] - a[3]);
    const aggregatedLength = processedData.length;

    // Add individual campaign data that meets the threshold
    data.forEach(row => {
        const impressions = parseInt(row['metrics.impressions']) || 0;
        if (impressions >= minPlaceImpr) {
            const cost = (Number(row['metrics.costMicros']) || 0) / 1e6;
            const interactions = Number(row['metrics.interactions']) || 0;
            const conversions = Number(row['metrics.conversions']) || 0;
            const value = Number(row['metrics.conversionsValue']) || 0;
            processedData.push([
                row['campaign.name'],
                row['detailPlacementView.groupPlacementTargetUrl'],
                row['detailPlacementView.placementType'],
                impressions,
                interactions,
                Number(row['metrics.videoViews']) || 0,
                cost,
                conversions,
                value,
                impressions > 0 ? interactions / impressions : 0,
                interactions > 0 ? conversions / interactions : 0,
                conversions > 0 ? cost / conversions : 0,
                cost > 0 ? value / cost : 0
            ]);
        }
    });

    // Sort individual campaign section
    const campaignData = processedData.slice(aggregatedLength);
    campaignData.sort((a, b) => {
        if (a[0] !== b[0]) return a[0].localeCompare(b[0]); // First by campaign name
        return b[3] - a[3]; // Then by impressions descending
    });

    // Reconstruct the array
    processedData = [...headers, ...processedData.slice(0, aggregatedLength), ...campaignData];

    return processedData;
}

function processShoppingProductsData(data) {
    if (!data || data.length === 0) return [['No Shopping Product data available']];
    const headers = ['Product ID', 'Title', 'Brand', 'Price', 'Condition', 'Availability', 'Impr.', 'Clicks', 'Cost', 'Conv.', 'Value', 'CTR', 'CvR', 'CPA', 'ROAS'];
    const rows = data.map(row => {
        const cost = (row['metrics.costMicros'] || 0) / 1e6;
        const clicks = row['metrics.clicks'] || 0;
        const impressions = row['metrics.impressions'] || 0;
        const conversions = row['metrics.conversions'] || 0;
        const value = row['metrics.conversionsValue'] || 0;
        return [
            row['shoppingProduct.itemId'],
            row['shoppingProduct.title'],
            row['shoppingProduct.brand'],
            (row['shoppingProduct.priceMicros'] || 0) / 1e6,
            row['shoppingProduct.condition'],
            row['shoppingProduct.availability'],
            impressions,
            clicks,
            cost,
            conversions,
            value,
            impressions > 0 ? clicks / impressions : 0,
            clicks > 0 ? conversions / clicks : 0,
            conversions > 0 ? cost / conversions : 0,
            cost > 0 ? value / cost : 0
        ];
    });
    return [headers, ...rows];
}

function processPmaxPlacementData(data, s) {
    if (!data || data.length === 0) {
        Logger.log('No PMax placement data available.');
        return [['No data available']];
    }

    let minPlaceImpr = s.minPlaceImpr;

    const totalImpressions = data.reduce((sum, row) => sum + (Number(row['metrics.impressions']) || 0), 0);

    const aggregatedData = new Map();

    data.forEach(row => {
        const key = JSON.stringify({
            displayName: row['performanceMaxPlacementView.displayName'],
            placement: row['performanceMaxPlacementView.placement'],
            type: row['performanceMaxPlacementView.placementType'],
            targetUrl: row['performanceMaxPlacementView.targetUrl']
        });

        if (!aggregatedData.has(key)) {
            aggregatedData.set(key, {
                displayName: row['performanceMaxPlacementView.displayName'],
                placement: row['performanceMaxPlacementView.placement'],
                type: row['performanceMaxPlacementView.placementType'],
                targetUrl: row['performanceMaxPlacementView.placementType'] === 'GOOGLE_PRODUCTS' ?
                    'Google Owned & Operated' :
                    row['performanceMaxPlacementView.displayName'] === 'Video no longer available' ?
                        'Video no longer available' : row['performanceMaxPlacementView.targetUrl'],
                impressions: 0
            });
        }

        aggregatedData.get(key).impressions += Number(row['metrics.impressions']) || 0;
    });

    let processedData = []
    const headers = [['Display Name', 'Placement', 'Type', 'Target URL', 'Campaign', 'Impressions']];

    for (const placementData of aggregatedData.values()) {
        if (placementData.impressions >= minPlaceImpr) {
            processedData.push([
                placementData.displayName,
                placementData.placement,
                placementData.type,
                placementData.targetUrl,
                'All Campaigns',
                placementData.impressions
            ]);
        }
    }

    const aggregatedLength = processedData.length;
    processedData.sort((a, b) => b[5] - a[5]);

    data.forEach(row => {
        const impressions = parseInt(row['metrics.impressions']) || 0;
        if (impressions >= minPlaceImpr) {
            const targetUrl = row['performanceMaxPlacementView.placementType'] === 'GOOGLE_PRODUCTS' ?
                'Google Owned & Operated' :
                row['performanceMaxPlacementView.displayName'] === 'Video no longer available' ?
                    'Video no longer available' : row['performanceMaxPlacementView.targetUrl'];

            processedData.push([
                row['performanceMaxPlacementView.displayName'],
                row['performanceMaxPlacementView.placement'],
                row['performanceMaxPlacementView.placementType'],
                targetUrl,
                row['campaign.name'],
                impressions
            ]);
        }
    });

    const campaignData = processedData.slice(aggregatedLength);
    campaignData.sort((a, b) => {
        if (a[4] !== b[4]) return a[4].localeCompare(b[4]);
        return b[5] - a[5];
    });

    processedData = [...headers, ...processedData.slice(0, aggregatedLength), ...campaignData];

    const campaigns = [...new Set(data.map(row => row['campaign.name']))].sort();
    processedData[0].push('All Campaigns', 'Total Impressions');
    for (let i = 1; i < processedData.length; i++) {
        processedData[i].push(i <= campaigns.length ? campaigns[i - 1] : '', '');
    }

    processedData[0][7] = totalImpressions;

    return processedData;
}

function extractAndFilterData(data) {
    let displayAssetsSet = new Set();
    let videoAssetsSet = new Set();
    let headlineAssetsSet = new Set();
    let descriptionAssetsSet = new Set();
    let longHeadlineAssetsSet = new Set();
    let logoAssetsSet = new Set();

    data.assetGroupAssetData.forEach(row => {
        if (row['assetGroupAsset.fieldType'] && row['assetGroupAsset.fieldType'].includes('MARKETING')) {
            displayAssetsSet.add(row['asset.resourceName']);
        }
        if (row['assetGroupAsset.fieldType'] && row['assetGroupAsset.fieldType'].includes('VIDEO')) {
            videoAssetsSet.add(row['asset.resourceName']);
        }
        if (row['assetGroupAsset.fieldType'] && row['assetGroupAsset.fieldType'].includes('HEADLINE')) {
            headlineAssetsSet.add(row['asset.resourceName']);
        }
        if (row['assetGroupAsset.fieldType'] && row['assetGroupAsset.fieldType'].includes('DESCRIPTION')) {
            descriptionAssetsSet.add(row['asset.resourceName']);
        }
        if (row['assetGroupAsset.fieldType'] && row['assetGroupAsset.fieldType'].includes('LONG_HEADLINE')) {
            longHeadlineAssetsSet.add(row['asset.resourceName']);
        }
        if (row['assetGroupAsset.fieldType'] && row['assetGroupAsset.fieldType'].includes('LOGO')) {
            logoAssetsSet.add(row['asset.resourceName']);
        }
    });

    let displayAssets = [...displayAssetsSet];
    let videoAssets = [...videoAssetsSet];
    let headlineAssets = [...headlineAssetsSet];
    let descriptionAssets = [...descriptionAssetsSet];
    let longHeadlineAssets = [...longHeadlineAssetsSet];
    let logoAssets = [...logoAssetsSet];

    let displayAssetData = data.displayVideoData.filter(row => displayAssets.includes(row['segments.assetInteractionTarget.asset']));
    let videoAssetData = data.displayVideoData.filter(row => videoAssets.includes(row['segments.assetInteractionTarget.asset']));
    let headlineAssetData = data.displayVideoData.filter(row => headlineAssets.includes(row['segments.assetInteractionTarget.asset']));
    let descriptionAssetData = data.displayVideoData.filter(row => descriptionAssets.includes(row['segments.assetInteractionTarget.asset']));
    let longHeadlineAssetData = data.displayVideoData.filter(row => longHeadlineAssets.includes(row['segments.assetInteractionTarget.asset']));
    let logoAssetData = data.displayVideoData.filter(row => logoAssets.includes(row['segments.assetInteractionTarget.asset']));

    return {
        displayAssetData,
        videoAssetData,
        headlineAssetData,
        descriptionAssetData,
        longHeadlineAssetData,
        logoAssetData
    };
}

function mergeAssetsWithDetails(displayMetrics, displayDetails, videoMetrics, videoDetails, headlineMetrics, headlineDetails, descriptionMetrics, descriptionDetails, longHeadlineMetrics, longHeadlineDetails, logoMetrics, logoDetails) {
    let mergeWithBucket = (metrics, details, bucket) => {
        return details.map(detail => {
            let metric = metrics[detail.assetName];
            return {
                ...detail,
                ...metric,
                bucket: bucket
            };
        });
    };

    let mergedDisplay = mergeWithBucket(displayMetrics, displayDetails, 'display');
    let mergedVideo = mergeWithBucket(videoMetrics, videoDetails, 'video');
    let mergedHeadline = mergeWithBucket(headlineMetrics, headlineDetails, 'headline');
    let mergedDescription = mergeWithBucket(descriptionMetrics, descriptionDetails, 'description');
    let mergedLongHeadline = mergeWithBucket(longHeadlineMetrics, longHeadlineDetails, 'long_headline');
    let mergedLogo = mergeWithBucket(logoMetrics, logoDetails, 'logo');

    let headers = ['Asset Name', 'Source', 'File/Title', 'URL/ID', 'Impr', 'Clicks', 'Views', 'Cost', 'Conv', 'Value', 'CTR', 'CvR', 'AOV', 'ROAS', 'CPA', 'Bucket'];
    let dataArray = [...mergedDisplay, ...mergedVideo, ...mergedHeadline, ...mergedDescription, ...mergedLongHeadline, ...mergedLogo].map(i => {
        return [i.assetName, i.assetSource, i.filenameOrTitle, i.urlOrID, i.impr, i.clicks, i.views, i.cost, i.conv, i.value, i.ctr, i.cvr, i.aov, i.roas, i.cpa, i.bucket];
    });

    return [headers].concat(dataArray);
}

function enrichAssetMetrics(aggregatedMetrics, assetData, type) {
    let assetDetailsArray = [];

    // For each asset in aggregatedMetrics, fetch details from assetData
    for (let assetName of Object.keys(aggregatedMetrics)) {
        // Find the asset in assetData
        let matchingAsset = assetData.find(asset => asset['asset.resourceName'] === assetName);
        if (matchingAsset) {
            let assetDetails = {
                type: type,
                assetName: assetName,
                assetSource: matchingAsset['asset.source'],
                filenameOrTitle: '',
                urlOrID: '',
                impr: aggregatedMetrics[assetName].impr
            };

            // Set filenameOrTitle and urlOrID based on asset type
            if (['display', 'logo'].includes(type)) {
                assetDetails.filenameOrTitle = matchingAsset['asset.name'] || matchingAsset['asset.imageAsset.fullSize.url'] || '';
                assetDetails.urlOrID = matchingAsset['asset.imageAsset.fullSize.url'] || '';
            } else if (type === 'video') {
                assetDetails.filenameOrTitle = matchingAsset['asset.youtubeVideoAsset.youtubeVideoTitle'] || '';
                assetDetails.urlOrID = matchingAsset['asset.youtubeVideoAsset.youtubeVideoId'] || '';
            } else if (['headline', 'description', 'long_headline'].includes(type)) {
                assetDetails.filenameOrTitle = matchingAsset['asset.textAsset.text'] || '';
                assetDetails.urlOrID = matchingAsset['asset.id'] || '';
            }

            assetDetailsArray.push(assetDetails);
        }
    }
    // sort by impr
    assetDetailsArray.sort((a, b) => b.impr - a.impr);
    return assetDetailsArray;
}

function aggregateMetricsByAsset(data) {
    const metrics = ['impr', 'clicks', 'views', 'cost', 'conv', 'value'];
    const calculations = {
        ctr: (m) => m.impr > 0 ? m.clicks / m.impr : 0,
        cvr: (m) => m.clicks > 0 ? m.conv / m.clicks : 0,
        aov: (m) => m.conv > 0 ? m.value / m.conv : 0,
        roas: (m) => m.cost > 0 ? m.value / m.cost : 0,
        cpa: (m) => m.conv > 0 ? m.cost / m.conv : 0
    };

    return data.reduce((acc, row) => {
        const asset = row['segments.assetInteractionTarget.asset'];
        if (!acc[asset]) {
            acc[asset] = metrics.reduce((m, k) => ({ ...m, [k]: 0 }), {});
        }

        // Update base metrics
        acc[asset].impr += parseInt(row['metrics.impressions']) || 0;
        acc[asset].clicks += parseInt(row['metrics.clicks']) || 0;
        acc[asset].views += parseInt(row['metrics.videoViews']) || 0;
        acc[asset].cost += (parseInt(row['metrics.costMicros']) / 1e6) || 0;
        acc[asset].conv += parseFloat(row['metrics.conversions']) || 0;
        acc[asset].value += parseFloat(row['metrics.conversionsValue']) || 0;

        // Calculate derived metrics
        Object.entries(calculations).forEach(([key, calc]) => {
            acc[asset][key] = calc(acc[asset]);
        });

        return acc;
    }, {});
}

function aggregateDataByDateAndCampaign(data) {
    let aggData = {};
    data.forEach(row => {
        let date = row['segments.date'];
        let campName = row['campaign.name'];
        let key = `${date}_${campName}`;
        if (!aggData[key]) {
            aggData[key] = {
                'date': date,
                'campName': campName,
                'cost': 0,
                'impr': 0,
                'clicks': 0,
                'conv': 0,
                'value': 0
            };
        }
        aggData[key].cost += row['metrics.costMicros'] ? row['metrics.costMicros'] / 1e6 : 0;
        aggData[key].impr += row['metrics.impressions'] ? parseInt(row['metrics.impressions']) : 0;
        aggData[key].clicks += row['metrics.clicks'] ? parseInt(row['metrics.clicks']) : 0;
        aggData[key].conv += row['metrics.conversions'] ? parseFloat(row['metrics.conversions']) : 0;
        aggData[key].value += row['metrics.conversionsValue'] ? parseFloat(row['metrics.conversionsValue']) : 0;
    });
    return Object.values(aggData);
}

function getSearchResults(processedCampData, nonSearchData) {
    let searchMetrics = {};

    // Pre-allocate all metrics to avoid dynamic object growth
    processedCampData.forEach(row => {
        let key = row.date + '_' + row.campName;
        searchMetrics[key] = {
            cost: 0, impr: 0, clicks: 0, conv: 0, value: 0
        };
    });

    // Single pass accumulation
    for (let i = 0; i < nonSearchData.length; i++) {
        let row = nonSearchData[i];
        let key = row.date + '_' + row.campName;
        let metrics = searchMetrics[key];
        if (metrics) {
            metrics.cost += row.cost || 0;
            metrics.impr += row.impressions || 0;
            metrics.clicks += row.clicks || 0;
            metrics.conv += row.conversions || 0;
            metrics.value += row.conversionsValue || 0;
        }
    }

    // Direct array allocation and population
    let results = new Array(processedCampData.length);
    for (let i = 0; i < processedCampData.length; i++) {
        let row = processedCampData[i];
        let key = row.date + '_' + row.campName;
        let nonSearch = searchMetrics[key];

        results[i] = {
            date: row.date,
            campName: row.campName,
            campType: row.campType,
            cost: row.cost - nonSearch.cost,
            impressions: row.impressions - nonSearch.impr,
            clicks: row.clicks - nonSearch.clicks,
            conversions: row.conversions - nonSearch.conv,
            conversionsValue: row.value - nonSearch.value
        };

        // Clean up negative values in one pass
        if (results[i].cost < 0) results[i].cost = 0;
        if (results[i].impressions < 0) results[i].impressions = 0;
        if (results[i].clicks < 0) results[i].clicks = 0;
        if (results[i].conversions < 0) results[i].conversions = 0;
        if (results[i].conversionsValue < 0) results[i].conversionsValue = 0;
    }

    return results;
}

function processData(data) {
    let summedData = {};
    data.forEach(row => {
        let date = row['segments.date'];
        let campName = row['campaign.name'];
        let campType = row['campaign.advertisingChannelType'];
        let key = `${date}_${campName}`;
        // Initialize if the key doesn't exist
        if (!summedData[key]) {
            summedData[key] = {
                'date': date,
                'campName': campName,
                'campType': campType,
                'cost': 0,
                'impr': 0,
                'clicks': 0,
                'conv': 0,
                'value': 0
            };
        }
        summedData[key].cost += row['metrics.costMicros'] ? row['metrics.costMicros'] / 1e6 : 0;
        summedData[key].impr += row['metrics.impressions'] ? parseInt(row['metrics.impressions']) : 0;
        summedData[key].clicks += row['metrics.clicks'] ? parseInt(row['metrics.clicks']) : 0;
        summedData[key].conv += row['metrics.conversions'] ? parseFloat(row['metrics.conversions']) : 0;
        summedData[key].value += row['metrics.conversionsValue'] ? parseFloat(row['metrics.conversionsValue']) : 0;
    });

    return Object.values(summedData);
}

function tidyAssetGroupData(data) {
    const calcMetrics = row => {
        const base = {
            impr: parseInt(row['metrics.impressions']) || 0,
            clicks: parseInt(row['metrics.clicks']) || 0,
            cost: (parseInt(row['metrics.costMicros']) / 1e6) || 0,
            conv: parseFloat(row['metrics.conversions']) || 0,
            value: parseFloat(row['metrics.conversionsValue']) || 0
        };

        return {
            'Camp Name': row['campaign.name'],
            'Asset Group Name': row['assetGroup.name'],
            'Status': row['assetGroup.status'],
            'Impr': base.impr,
            'Clicks': base.clicks,
            'Cost': base.cost,
            'Conv': base.conv,
            'Value': base.value,
            'CTR': base.impr > 0 ? base.clicks / base.impr : 0,
            'CVR': base.clicks > 0 ? base.conv / base.clicks : 0,
            'AOV': base.conv > 0 ? base.value / base.conv : 0,
            'ROAS': base.cost > 0 ? base.value / base.cost : 0,
            'CPA': base.conv > 0 ? base.cost / base.conv : 0,
            'Strength': row['assetGroup.adStrength']
        };
    };

    return data
        .map(calcMetrics)
        .sort((a, b) => a['Camp Name'].localeCompare(b['Camp Name']) || b['Impr'] - a['Impr']);
}

function processDataCommon(dataGroups, isSummary) {
    // Pre-allocate arrays for headers
    let headerRow = isSummary ?
        ['Date', 'Campaign Name', 'Camp Cost', 'Camp Conv', 'Camp Value',
            'Shop Cost', 'Shop Conv', 'Shop Value', 'Disp Cost', 'Disp Conv', 'Disp Value',
            'Video Cost', 'Video Conv', 'Video Value', 'Search Cost', 'Search Conv', 'Search Value',
            'Campaign Type'] :
        ['Campaign Name', 'Camp Cost', 'Camp Conv', 'Camp Value', 'Shop Cost', 'Shop Conv',
            'Shop Value', 'Disp Cost', 'Disp Conv', 'Disp Value', 'Video Cost', 'Video Conv',
            'Video Value', 'Search Cost', 'Search Conv', 'Search Value', 'Campaign Type'];

    let processed = {};
    let keyOrder = [];

    // Pre-process to get unique keys and initialize data structure
    dataGroups.forEach(group => {
        group.data.forEach(row => {
            let key = isSummary ? row.date + '_' + row.campName : row.campName;
            if (!processed[key]) {
                processed[key] = {
                    date: row.date || '',
                    campName: row.campName,
                    campType: row.campType,
                    general: [0, 0, 0],
                    shopping: [0, 0, 0],
                    display: [0, 0, 0],
                    video: [0, 0, 0],
                    search: [0, 0, 0]
                };
                keyOrder.push(key);
            }
        });
    });

    // Process data in direct assignments
    for (let i = 0; i < dataGroups.length; i++) {
        let group = dataGroups[i];
        let data = group.data;
        let type = group.type;

        for (let j = 0; j < data.length; j++) {
            let row = data[j];
            let key = isSummary ? row.date + '_' + row.campName : row.campName;
            let entry = processed[key];

            let targetArray = entry[type];
            targetArray[0] += row.cost || 0;
            targetArray[1] += row.conv || 0;
            targetArray[2] += row.value || 0;

            if (type === 'general' && row.campType === 'SHOPPING') {
                let shopArray = entry.shopping;
                shopArray[0] += row.cost || 0;
                shopArray[1] += row.conv || 0;
                shopArray[2] += row.value || 0;
            }
        }
    }

    // Build output array directly
    let output = new Array(keyOrder.length + 1);
    output[0] = headerRow;

    for (let i = 0; i < keyOrder.length; i++) {
        let data = processed[keyOrder[i]];
        let remainingCost;

        // Adjust video
        remainingCost = data.general[0] - data.shopping[0];
        if (data.video[0] > remainingCost) {
            data.video[0] = remainingCost;
            data.video[1] = Math.min(data.video[1], data.general[1] - data.shopping[1]);
            data.video[2] = Math.min(data.video[2], data.general[2] - data.shopping[2]);
        }

        // Adjust display
        remainingCost = data.general[0] - data.shopping[0] - data.video[0];
        if (data.display[0] > remainingCost) {
            data.display[0] = remainingCost;
            data.display[1] = Math.min(data.display[1], data.general[1] - data.shopping[1] - data.video[1]);
            data.display[2] = Math.min(data.display[2], data.general[2] - data.shopping[2] - data.video[2]);
        }

        // Calculate search
        remainingCost = data.general[0] - data.shopping[0] - data.video[0] - data.display[0];
        if (remainingCost <= 0) {
            data.search = [0, 0, 0];
        } else {
            data.search[0] = remainingCost;
            data.search[1] = Math.max(0, data.general[1] - data.shopping[1] - data.video[1] - data.display[1]);
            data.search[2] = Math.max(0, data.general[2] - data.shopping[2] - data.video[2] - data.display[2]);
        }

        // Build row directly
        let row = isSummary ?
            [data.date, data.campName].concat(
                data.general, data.shopping, data.display,
                data.video, data.search, [data.campType]
            ) :
            [data.campName].concat(
                data.general, data.shopping, data.display,
                data.video, data.search, [data.campType]
            );

        output[i + 1] = row;
    }

    return output;
}

function processSummaryData(processedCampData, processedAssetGroupData, processedDisplayData, processedVideoData, searchResults) {
    let dataGroups = [
        { data: processedCampData, type: 'general' },
        { data: processedAssetGroupData, type: 'shopping', excludeType: 'SHOPPING' },
        { data: processedDisplayData, type: 'display' },
        { data: processedVideoData, type: 'video' },
        { data: searchResults, type: 'search' }
    ];
    return processDataCommon(dataGroups, true);
}

function processTotalData(processedCampData, processedAssetGroupData, processedDisplayData, processedVideoData, searchResults) {
    let dataGroups = [
        { data: processedCampData, type: 'general' },
        { data: processedAssetGroupData, type: 'shopping', excludeType: 'SHOPPING' },
        { data: processedDisplayData, type: 'display' },
        { data: processedVideoData, type: 'video' },
        { data: searchResults, type: 'search' }
    ];
    return processDataCommon(dataGroups, false);
}

function fetchProductData(queryString, tCost, tRoas, outputType) {
    let productCampaigns = new Map();
    let aggregatedData = new Map();
    let iterator = AdsApp.search(queryString);

    while (iterator.hasNext()) {
        let row = flattenObject(iterator.next());
        let productId = row['segments.productItemId'];
        let channelType = row['campaign.advertisingChannelType'];
        let cost = (Number(row['metrics.costMicros']) / 1e6) || 0;

        // Track campaigns and costs
        if (!productCampaigns.has(productId)) {
            productCampaigns.set(productId, {
                campaigns: { SHOPPING: new Set(), PERFORMANCE_MAX: new Set() },
                costs: { SHOPPING: 0, PERFORMANCE_MAX: 0 }
            });
        }

        if (['SHOPPING', 'PERFORMANCE_MAX'].includes(channelType)) {
            let info = productCampaigns.get(productId);
            info.campaigns[channelType].add(row['campaign.name']);
            info.costs[channelType] += cost;
        }

        // Aggregate metrics
        let key = getUniqueKey(row, outputType);
        if (!aggregatedData.has(key)) {
            aggregatedData.set(key, {
                'Impr': 0, 'Clicks': 0, 'Cost': 0, 'Conv': 0, 'Value': 0,
                'Product Title': row['segments.productTitle'],
                'Product ID': row['segments.productItemId'],
                'Campaign': row['campaign.name'],
                'Channel': row['campaign.advertisingChannelType']
            });
        }

        let metrics = aggregatedData.get(key);
        metrics.Impr += Number(row['metrics.impressions']) || 0;
        metrics.Clicks += Number(row['metrics.clicks']) || 0;
        metrics.Cost += cost;
        metrics.Conv += Number(row['metrics.conversions']) || 0;
        metrics.Value += Number(row['metrics.conversionsValue']) || 0;
    }

    // Transform to final output using array method
    return [...aggregatedData.values()].map(data => {
        let campaignInfo = productCampaigns.get(data['Product ID']);
        let baseObj = {
            ...data,
            countPmax: campaignInfo?.campaigns.PERFORMANCE_MAX.size || 0,
            countShop: campaignInfo?.campaigns.SHOPPING.size || 0,
            costPmax: campaignInfo?.costs.PERFORMANCE_MAX || 0,
            costShop: campaignInfo?.costs.SHOPPING || 0,
            inBoth: campaignInfo?.campaigns.PERFORMANCE_MAX.size &&
                campaignInfo?.campaigns.SHOPPING.size ? 1 : 0
        };

        if (outputType !== 'idCount') {
            baseObj.ROAS = baseObj.Cost > 0 ? baseObj.Value / baseObj.Cost : 0;
            baseObj.CvR = baseObj.Clicks > 0 ? baseObj.Conv / baseObj.Clicks : 0;
            baseObj.CTR = baseObj.Impr > 0 ? baseObj.Clicks / baseObj.Impr : 0;
            baseObj.Bucket = determineBucket(baseObj.Cost, baseObj.Conv, baseObj.ROAS, tCost, tRoas);
        }

        return constructBaseDataObject(baseObj, outputType);
    });
}

function getUniqueKey(row, type) {
    let id = row['segments.productItemId'];
    let title = row['segments.productTitle'];
    let campaign = row['campaign.name'];
    let channel = row['campaign.advertisingChannelType'];

    switch (type) {
        case 'pTitle': return title;
        case 'pTitleCampaign': return title + '|' + campaign;
        case 'pTitleID': return title + '|' + id + '|' + campaign;
        case 'idAccount': return id;
        case 'idChannel': return id + '|' + channel;
        case 'idCount': return id;
        default: return id;
    }
}

function determineBucket(cost, conv, roas, tCost, tRoas) {
    if (cost === 0) return 'zombie';
    if (conv === 0) return 'zeroconv';
    if (cost < tCost) return roas < tRoas ? 'meh' : 'flukes';
    return roas < tRoas ? 'costly' : 'profitable';
}

function constructBaseDataObject(aggData, outputType) {
    let baseDataObject = {
        'Product Title': aggData['Product Title'],
        'Product ID': aggData['Product ID'],
        'Impr': aggData['Impr'],
        'Clicks': aggData['Clicks'],
        'Cost': aggData['Cost'],
        'Conv': aggData['Conv'],
        'Value': aggData['Value'],
        'CTR': aggData['CTR'],
        'ROAS': aggData['ROAS'],
        'CvR': aggData['CvR'],
        'Bucket': aggData['Bucket'],
        'Campaign': aggData['Campaign'],
        'Channel': aggData['Channel']
    };

    switch (outputType) {
        // build with all columns - which is pTitle&ID
        case 'pTitleCampaign': // Remove the Product ID for 'pTitleCampaign' output type
            delete baseDataObject['Product ID'];
            break;
        case 'pTitle': // delete ID, campaign & channel - to aggregate across whole account
            delete baseDataObject['Product ID'];
            delete baseDataObject['Campaign'];
            delete baseDataObject['Channel'];
            break;

        // Account level is default (v60) & channel if 'idChannel' is selected
        case 'idAccount':
            baseDataObject = {
                'ID': aggData['Product ID'],
                'Bucket': aggData['Bucket']
            };
            break;
        case 'idChannel':
            baseDataObject = {
                'ID': aggData['Product ID'],
                'Bucket': aggData['Bucket'],
                'Channel': aggData['Channel']
            };
            break;
        case 'idCount':
            baseDataObject = {
                'Product Title': aggData['Product Title'],
                'ID': aggData['Product ID'],
                'inBoth': aggData['inBoth'] || 0,
                'countPmax': aggData['countPmax'] || 0,
                'countShop': aggData['countShop'] || 0,
                'costPmax': aggData['costPmax'] || 0,
                'costShop': aggData['costShop'] || 0
            }
            break;
    }

    return baseDataObject;
}

function extractAndAggregateTitleNGrams(s, productData) {
    let nGrams = {};

    productData.forEach(row => {
        let productTitle = row['Product Title'] ? row['Product Title'].toLowerCase() : '';
        let terms = productTitle.split(' ');

        terms.forEach((term) => {
            let key = cleanNGram(term);
            if (!nGrams[key]) {
                nGrams[key] = {
                    nGram: term,
                    impr: 0,
                    clicks: 0,
                    cost: 0,
                    conv: 0,
                    value: 0
                };
            }

            nGrams[key].impr += row['Impr'];
            nGrams[key].clicks += row['Clicks'];
            nGrams[key].cost += row['Cost'];
            nGrams[key].conv += row['Conv'];
            nGrams[key].value += row['Value'];
        });
    });

    let tNgrams = [['nGram', 'Impr', 'Clicks', 'Cost', 'Conv', 'Value', 'CTR', 'CvR', 'AOV', 'ROAS']];

    for (let term in nGrams) {
        let item = nGrams[term];
        item.CTR = item.impr > 0 ? item.clicks / item.impr : 0;
        item.CvR = item.clicks > 0 ? item.conv / item.clicks : 0;
        item.AOV = item.conv > 0 ? item.value / item.conv : 0;
        item.ROAS = item.cost > 0 ? item.value / item.cost : 0;
        tNgrams.push([item.nGram, item.impr, item.clicks, item.cost, item.conv, item.value, item.CTR, item.CvR, item.AOV, item.ROAS]);
    }

    tNgrams.sort((a, b) => {
        if (a[0] === 'nGram') return -1;
        if (b[0] === 'nGram') return 1;
        return b[3] - a[3]; // Sort by Cost (index 3) in descending order
    });

    let brandTerm = s.brandTerm.includes(',') ? s.brandTerm.split(/[ ,]+/).map(i => i.toLowerCase()) : s.brandTerm.split(' ').map(i => i.toLowerCase());
    tNgrams = tNgrams.filter(i => !brandTerm.includes(i[0]) && i[0].length > 1);

    return tNgrams;
}

function cleanNGram(nGram) {
    if (!nGram || nGram.length <= 1) return '';

    let start = 0;
    let end = nGram.length;
    let specialChars = '.,/#!$%^&*;:{}=-_`~()';

    // Trim start
    while (start < end && specialChars.indexOf(nGram[start]) > -1) start++;

    // Trim end
    while (end > start && specialChars.indexOf(nGram[end - 1]) > -1) end--;

    return end - start <= 1 ? '' : nGram.slice(start, end);
}

function log(ss, startDuration, s, ident) {
    let endDuration = new Date();
    let duration = ((endDuration - startDuration) / 1000).toFixed(0);
    Logger.log(`Script execution time: ${duration} seconds. \nFinished script for ${ident}.`);

    let newRow = [new Date(), duration, s.numberOfDays, s.tCost, s.tRoas, s.brandTerm, ident, s.aiRunAt,
    s.fromDate, s.toDate, s.lotsProducts, s.turnonTitleCampaign, s.turnonTitleID, s.turnonIDChannel,
    s.campFilter, s.turnonLP, s.turnonPlace, s.turnonGeo, s.turnonChange, s.turnonAISheet, s.turnonNgramSheet,
    s.llm, s.model, s.lang, s.scriptVersion, s.sheetVersion, s.turnonDemandGen];
    logUrl = ss.getRangeByName('u').getValue();
    [safeOpenAndShareSpreadsheet(logUrl), ss].map(s => s.getSheetByName('log')).forEach(sheet => sheet.appendRow(newRow));
}

function determineBucketAndDistance(term, brandTerms, tolerance) {
    const terms = brandTerms?.split(',').map(t => t.trim()).filter(Boolean) || ['no brand has been entered on the sheet'];
    if (!term || term === 'blank') return { bucket: 'blank', distance: '' };

    const exactMatch = terms.some(b => term === b.toLowerCase());
    if (exactMatch) return { bucket: 'brand', distance: 0 };

    const closeMatch = terms.reduce((match, brand) => {
        const dist = levenshtein(term, brand.toLowerCase());
        return (!match && (term.includes(brand.toLowerCase()) || dist <= tolerance)) ?
            { bucket: 'close-brand', distance: dist } : match;
    }, null);

    return closeMatch || { bucket: 'non-brand', distance: '' };
}

function levenshtein(a, b) {
    let tmp;
    if (a.length === 0) { return b.length; }
    if (b.length === 0) { return a.length; }
    if (a.length > b.length) { tmp = a; a = b; b = tmp; }

    let i, j, res, alen = a.length, blen = b.length, row = Array(alen);
    for (i = 0; i <= alen; i++) { row[i] = i; }

    for (i = 1; i <= blen; i++) {
        res = i;
        for (j = 1; j <= alen; j++) {
            tmp = row[j - 1];
            row[j - 1] = res;
            res = b[i - 1] === a[j - 1] ? tmp : Math.min(tmp + 1, Math.min(res + 1, row[j] + 1));
        }
    }
    return res;
}

function aggLPData(lpData, ss, s, tabName) {
    let urlData = {};

    lpData.forEach(row => {
        let url = row['landingPageView.unexpandedFinalUrl'];
        if (url) {
            let [domain, ...pathSegments] = url.split('/').slice(2);
            let paths = [domain, ...pathSegments.join('/').split('?')[0].split('/')];

            paths.forEach((path, i) => {
                path = path || '';
                let key = `/${paths.slice(1, i + 1).join('/')}`;
                if (!urlData[key]) {
                    urlData[key] = {
                        domain: paths[0],
                        path1: paths[1] || '',
                        path2: paths[2] || '',
                        path3: paths[3] || '',
                        impr: 0,
                        clicks: 0,
                        cost: 0,
                        conv: 0,
                        value: 0,
                        bucket: ['domain', 'path1', 'path2', 'path3'][i]
                    };
                }
                urlData[key].impr += parseFloat(row['metrics.impressions']) || 0;
                urlData[key].clicks += parseFloat(row['metrics.clicks']) || 0;
                urlData[key].cost += parseFloat(row['metrics.costMicros']) / 1e6 || 0;
                urlData[key].conv += parseFloat(row['metrics.conversions']) || 0;
                urlData[key].value += parseFloat(row['metrics.conversionsValue']) || 0;
            });
        }
    });

    // Calculate additional metrics
    Object.values(urlData).forEach(d => {
        d.ctr = d.impr ? d.clicks / d.impr : 0;
        d.cvr = d.clicks ? d.conv / d.clicks : 0;
        d.aov = d.conv ? d.value / d.conv : 0;
        d.roas = d.cost ? d.value / d.cost : 0;
        d.cpa = d.conv ? d.cost / d.conv : 0;
    });

    // Prepare data and output
    let urlSummary = Object.entries(urlData)
        .filter(([url, d]) => d.impr > s.minLpImpr)
        .map(([url, d]) => [
            url, d.impr, d.clicks, d.cost, d.conv, d.value, d.ctr, d.cvr, d.aov, d.roas, d.cpa, d.bucket,
            d.domain,
            d.path1,
            d.path2,
            d.path3
        ]);

    urlSummary.sort((a, b) => b[1] - a[1]);
    let urlHeaders = [['PathDetail', 'Impr', 'Clicks', 'Cost', 'Conv', 'Value', 'CTR', 'CVR', 'AOV', 'ROAS', 'CPA', 'Bucket', 'Domain', 'Path1', 'Path2', 'Path3']];
    outputDataToSheet(ss, tabName, urlHeaders.concat(urlSummary), 'notLast');
}

function updateCrossReferences(clientSheet, whispererSheet) {
    try {
        clientSheet.getRangeByName('whispererUrl').setValue(whispererSheet.getUrl());
        whispererSheet.getRangeByName('pmaxUrl').setValue(clientSheet.getUrl());
    } catch (e) {
        Logger.error('Error updating cross-references: ' + e.message);
    }
}
//#endregion

//#region MCC only functions 
function safeOpenAndShareSpreadsheet(url, setAccess = false, newName = null) {
    try {
        // Basic validation
        if (!url) {
            console.error(`URL is empty or undefined: ${url}`);
            return null;
        }

        // Type checking and format validation
        if (typeof url !== 'string') {
            console.error(`Invalid URL type - expected string but got ${typeof url}`);
            return null;
        }

        // Validate Google Sheets URL format
        if (!url.includes('docs.google.com/spreadsheets/d/')) {
            console.error(`Invalid Google Sheets URL format: ${url}`);
            return null;
        }

        // Try to open the spreadsheet
        let ss;
        try {
            ss = SpreadsheetApp.openByUrl(url);
        } catch (error) {
            Logger.log(`Error opening spreadsheet: ${error.message}`);
            Logger.log(`Settings were: ${url}, ${setAccess}, ${newName}`);
            return null;
        }

        // Handle copy if newName is provided
        if (newName) {
            try {
                ss = ss.copy(newName);
            } catch (error) {
                Logger.log(`Error copying spreadsheet: ${error.message}`);
                return null;
            }
        }

        // Handle sharing settings if required
        if (setAccess) {
            try {
                let file = DriveApp.getFileById(ss.getId());

                // Try ANYONE_WITH_LINK first
                try {
                    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
                    Logger.log("Sharing set to ANYONE_WITH_LINK");
                } catch (error) {
                    Logger.log("ANYONE_WITH_LINK failed, trying DOMAIN_WITH_LINK");

                    // If ANYONE_WITH_LINK fails, try DOMAIN_WITH_LINK
                    try {
                        file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
                        Logger.log("Sharing set to DOMAIN_WITH_LINK");
                    } catch (error) {
                        Logger.log("DOMAIN_WITH_LINK failed, setting to PRIVATE");

                        // If all else fails, set to PRIVATE
                        try {
                            file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);
                            Logger.log("Sharing set to PRIVATE");
                        } catch (error) {
                            Logger.log(`Failed to set any sharing permissions: ${error.message}`);
                        }
                    }
                }
            } catch (error) {
                Logger.log(`Error setting file permissions: ${error.message}`);
                // Continue even if sharing fails - the sheet is still usable
            }
        }

        return ss;

    } catch (error) {
        // Catch any other unexpected errors
        console.error(`Unexpected error in safeOpenAndShareSpreadsheet: ${error.message}`);
        Logger.log(`Full error details: ${error.stack}`);
        return null;
    }
}

//#endregion




/*

DISCLAIMER  -  PLEASE READ CAREFULLY BEFORE USING THIS SCRIPT

Fair Use: This script is provided for the sole use of the entity (business, agency, or individual) to which it is licensed.
While you are encouraged to use and benefit from this script, you must do so within the confines of this agreement.

Copyright: All rights, including copyright, in this script are owned by or licensed to Off Rhodes Pty Ltd t/a Mike Rhodes Ideas.
Reproducing, distributing, or selling any version of this script, whether modified or unmodified, without proper authorization is strictly prohibited.

License Requirement: A separate license must be purchased for each legal entity that wishes to use this script.
For example, if you own multiple businesses or agencies, each business or agency can use this script under one license.
However, if you are part of a holding group or conglomerate with multiple separate entities, each entity must purchase its own license for use.

Code of Honour: This script is offered under a code of honour.
We trust in the integrity of our users to adhere to the terms of this agreement and to do the right thing.
Your honour and professionalism in respecting these terms not only supports the creator but also fosters a community of trust and respect.

Limitations & Liabilities: Off Rhodes Pty Ltd t/a Mike Rhodes Ideas does not guarantee that this script will be error-free
or will meet your specific requirements. We are not responsible for any damages or losses that might arise from using this script.
Always back up your data and test the script in a safe environment before deploying it in a production setting.

The script does not make any changes to your account or data.
It only reads data from your account and writes it to your spreadsheet.
However if you choose to use the data on the ID tab in a supplemental data feed in your GMC account, you do so at your own risk.

By using this script, you acknowledge that you have read, understood, and agree to be bound by the terms of this license agreement.
If you do not agree with these terms, do not use this script.




Help Docs: https://pmax.super.site/



--------------------------------------------------------------------------------------------------------------
Please feel free to change the code... it's why I add so many comments. I'd love it if you get to know how it works.
However, if you do, it's because you've read the wiki & you know code, scripts and google ads inside out
and (importantly!) you're happy to no longer receive any support. I can't support code that's been changed, sorry.
-------------------------------------------------------------------------------------------------------------- */



// If you get any errors, please read the docs at https://pmax.super.site/ & then try again ;)

// If you're still getting an error, copy the logs & paste them into a post at https://mikerhodes.circle.so/c/help/ 
// and tag me so I can help you resolve it :)


// Now hit preview (or run) and let's get this party started! Thanks for using this script.


// PS you're awesome! 
