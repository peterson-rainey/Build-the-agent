// v2 pmax daily assets (c) MikeRhodes.com.au  
// https://www.loom.com/share/9002d07b63894b20baa567139fb5e9be?sid=68904029-de56-4eb9-bad6-ce2b2ae2cedb

// enter your URL here between the single quotes - copy the template sheet first : https://docs.google.com/spreadsheets/d/1TRuMopr2CIiy7vUPWMAPfr1xn1R6xZ3Zm_jPYKv6Jp4/copy
const SHEET_URL = '';


// change this if needed (tiny accounts can use 1, large accounts (>$50k/month) should probably use 1000 or more). You may need to experiment. If script takes ages to run, or sheet crashes, increase this.
const IMPRESSION_THRESHOLD = 100;




// please don't change any code below this line  ---------------------------------------------------------------------------------------------------------------------------


const PERCENTAGE = 0.8;
const TARGET_ID = '';  // use for debugging
const TAB_NAMES = {
    DAILY: 'dataDaily',
    WEEKLY: 'dataWeekly',
    MONTHLY: 'dataMonthly'
};

const TIME_PERIODS = [
    { name: 'DAILY', days: 30, segmentType: 'date' },
    { name: 'WEEKLY', days: 365, segmentType: 'week' }, // approx 52 weeks
    { name: 'MONTHLY', days: 1100, segmentType: 'month' }  // approx 36 months
];

function main() {
    let ss = SpreadsheetApp.openByUrl(SHEET_URL);

    for (let period of TIME_PERIODS) {
        try {
            Logger.log(`\n--- Processing ${period.name} data ---`);
            let dateRange = getLastNDays(period.days);
            Logger.log(`Date range: ${dateRange.start} to ${dateRange.end}`);

            let assetGroupAssetData = fetchAssetGroupAssetData();
            let campaignData = fetchCampaignData(dateRange, period.segmentType);
            let displayVideoData = fetchDisplayVideoData(dateRange, period.segmentType);
            let assetData = fetchAssetData();

            let pmaxAssets = identifyPMaxAssets(assetGroupAssetData);
            let campaignMetrics = aggregateCampaignData(campaignData);
            let assetPerformance = processAssetPerformance(displayVideoData, pmaxAssets, period.segmentType);
            let searchPerformance = calculateSearchPerformance(campaignMetrics, assetPerformance);

            let finalAssetData = combineAssetData(assetData, pmaxAssets, assetPerformance);
            writeDataToSheet(ss, finalAssetData, TAB_NAMES[period.name]);
            addNamedRanges(ss, TAB_NAMES[period.name], period.name);

            // Set initial camp values for this period
            setInitialCampValues(ss, period);

            Logger.log(`Successfully processed ${period.name} data`);
        } catch (error) {
            Logger.log(`Error processing ${period.name} data: ${error.message}`);
        }
    }
}

function getLastNDays(n) {
    n = n !== undefined ? n : 30;
    let endDate = new Date();
    let startDate = new Date(endDate.getTime());
    startDate.setDate(endDate.getDate() - n);
    endDate.setDate(endDate.getDate() - 1);

    let timeZone = AdsApp.currentAccount().getTimeZone();
    return {
        start: Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd'),
        end: Utilities.formatDate(endDate, timeZone, 'yyyy-MM-dd')
    };
}

function fetchAssetGroupAssetData() {
    let query = `
      SELECT 
        campaign.name,
        asset_group.name,
        asset_group_asset.field_type,
        asset.resource_name
      FROM asset_group_asset
      WHERE campaign.advertising_channel_type = "PERFORMANCE_MAX"
      AND asset_group_asset.field_type NOT IN ( "BUSINESS_NAME", "CALL_TO_ACTION_SELECTION")`;
    return executeQuery(query);
}

function fetchAssetData() {
    let query = `
      SELECT
        asset.resource_name,
        asset.name,
        asset.type,
        asset.image_asset.full_size.url,
        asset.text_asset.text,
        asset.youtube_video_asset.youtube_video_title,
        asset.youtube_video_asset.youtube_video_id
      FROM asset
      WHERE asset.type IN ('TEXT', 'IMAGE', 'YOUTUBE_VIDEO')`;
    return executeQuery(query);
}

function fetchCampaignData(dateRange, segmentType) {
    let query = `
      SELECT
        campaign.name,
        campaign.id,
        segments.date,
        metrics.cost_micros,
        metrics.conversions,
        metrics.conversions_value,
        metrics.impressions,
        metrics.clicks
      FROM campaign
      WHERE metrics.impressions > ${IMPRESSION_THRESHOLD}
        AND campaign.advertising_channel_type = "PERFORMANCE_MAX"
        AND segments.date BETWEEN "${dateRange.start}" AND "${dateRange.end}"`;
    return executeQuery(query);
}

function fetchDisplayVideoData(dateRange, segmentType) {
    let query = `
      SELECT
        campaign.name,
        segments.date,
        segments.asset_interaction_target.asset,
        segments.asset_interaction_target.interaction_on_this_asset,
        metrics.cost_micros,
        metrics.conversions,
        metrics.conversions_value,
        metrics.impressions,
        metrics.clicks,
        metrics.video_views
      FROM campaign
      WHERE metrics.impressions > ${IMPRESSION_THRESHOLD}
        AND campaign.advertising_channel_type = "PERFORMANCE_MAX"
        AND segments.date BETWEEN "${dateRange.start}" AND "${dateRange.end}"
        AND segments.asset_interaction_target.interaction_on_this_asset = FALSE`;
    return executeQuery(query);
}

function executeQuery(query) {
    let report = AdsApp.report(query);
    let rows = [];
    let reportRow = report.rows();
    while (reportRow.hasNext()) {
        rows.push(reportRow.next());
    }
    return rows;
}

function identifyPMaxAssets(assetGroupAssetData) {
    let pmaxAssets = {};
    assetGroupAssetData.forEach(row => {
        let assetId = row['asset.resource_name'];
        if (!pmaxAssets[assetId]) {
            pmaxAssets[assetId] = {
                campaigns: new Set(),
                fieldTypes: new Set()
            };
        }
        pmaxAssets[assetId].campaigns.add(row['campaign.name']);
        pmaxAssets[assetId].fieldTypes.add(row['asset_group_asset.field_type']);
    });
    return pmaxAssets;
}

function aggregateCampaignData(campaignData) {
    let aggregatedData = {};
    campaignData.forEach(row => {
        let campaignName = row['campaign.name'];
        if (!aggregatedData[campaignName]) {
            aggregatedData[campaignName] = {
                cost: 0,
                conversions: 0,
                conversionValue: 0,
                impressions: 0,
                clicks: 0
            };
        }
        aggregatedData[campaignName].cost += Number(row['metrics.cost_micros'] || 0) / 1000000;
        aggregatedData[campaignName].conversions += Number(row['metrics.conversions'] || 0);
        aggregatedData[campaignName].conversionValue += Number(row['metrics.conversions_value'] || 0);
        aggregatedData[campaignName].impressions += Number(row['metrics.impressions'] || 0);
        aggregatedData[campaignName].clicks += Number(row['metrics.clicks'] || 0);
    });
    return aggregatedData;
}

function calculateSearchPerformance(campaignMetrics, assetPerformance) {
    let searchPerformance = {};
    for (let campaign in campaignMetrics) {
        searchPerformance[campaign] = {
            cost: campaignMetrics[campaign].cost,
            conversions: campaignMetrics[campaign].conversions,
            conversionValue: campaignMetrics[campaign].conversionValue,
            impressions: campaignMetrics[campaign].impressions,
            clicks: campaignMetrics[campaign].clicks
        };
    }

    // Subtract asset performance from total campaign performance
    for (let assetId in assetPerformance) {
        for (let campaign in searchPerformance) {
            searchPerformance[campaign].cost -= assetPerformance[assetId].cost;
            searchPerformance[campaign].conversions -= assetPerformance[assetId].conversions;
            searchPerformance[campaign].conversionValue -= assetPerformance[assetId].conversionValue;
            searchPerformance[campaign].impressions -= assetPerformance[assetId].impressions;
            searchPerformance[campaign].clicks -= assetPerformance[assetId].clicks;
        }
    }

    return searchPerformance;
}

function processAssetPerformance(displayVideoData, pmaxAssets, segmentType) {
    let assetPerformance = {};

    displayVideoData.forEach(row => {
        let assetId = row['segments.asset_interaction_target.asset'];
        let date = row['segments.date'];
        let campaignName = row['campaign.name'];

        if (!pmaxAssets[assetId]) return; // Skip if not a PMax asset

        let segmentDate = getSegmentDate(date, segmentType);

        if (!assetPerformance[assetId]) {
            assetPerformance[assetId] = {};
        }

        if (!assetPerformance[assetId][segmentDate]) {
            assetPerformance[assetId][segmentDate] = {};
        }

        if (!assetPerformance[assetId][segmentDate][campaignName]) {
            assetPerformance[assetId][segmentDate][campaignName] = {
                impressions: 0,
                clicks: 0,
                cost: 0,
                conversions: 0,
                conversionValue: 0,
                videoViews: 0
            };
        }

        let metrics = assetPerformance[assetId][segmentDate][campaignName];
        metrics.impressions += Number(row['metrics.impressions'] || 0);
        metrics.clicks += Number(row['metrics.clicks'] || 0);
        metrics.cost += Number(row['metrics.cost_micros'] || 0) / 1000000;
        metrics.conversions += Number(row['metrics.conversions'] || 0);
        metrics.conversionValue += Number(row['metrics.conversions_value'] || 0);
        metrics.videoViews += Number(row['metrics.video_views'] || 0);
    });

    return assetPerformance;
}

function getSegmentDate(date, segmentType) {
    let d = new Date(date);
    switch (segmentType) {
        case 'date':
            return date; // Keep daily format
        case 'week':
            d.setDate(d.getDate() - d.getDay()); // Set to the start of the week (Sunday)
            return Utilities.formatDate(d, AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
        case 'month':
            return date.substring(0, 7); // Return YYYY-MM format
        default:
            return date;
    }
}

function getAssetContent(row) {
    switch (row['asset.type']) {
        case 'TEXT':
            return row['asset.text_asset.text'];
        case 'IMAGE':
            return row['asset.image_asset.full_size.url'];
        case 'YOUTUBE_VIDEO':
            let content = row['asset.youtube_video_asset.youtube_video_id'];
            return `https://www.youtube.com/watch?v=${content}`;
        default:
            return '';
    }
}

function select_highest_priority_field_type(field_types) {
    const priority_order = [
        "HEADLINE", "DESCRIPTION", "LONG_HEADLINE", "CALLOUT",
        "CALL_TO_ACTION_SELECTION", "BUSINESS_NAME", "SITELINK",
        "PROMOTION", "PRICE", "STRUCTURED_SNIPPET", "CALL",
        "LOGO", "BUSINESS_LOGO", "LANDSCAPE_LOGO", "AD_IMAGE",
        "MARKETING_IMAGE", "PORTRAIT_MARKETING_IMAGE", "SQUARE_MARKETING_IMAGE",
        "VIDEO", "YOUTUBE_VIDEO", "MEDIA_BUNDLE", "BOOK_ON_GOOGLE",
        "HOTEL_PROPERTY", "HOTEL_CALLOUT", "LEAD_FORM", "MOBILE_APP",
        "DEMAND_GEN_CAROUSEL_CARD", "MANDATORY_AD_TEXT", "UNKNOWN", "UNSPECIFIED"
    ];

    for (let field_type of priority_order) {
        if (field_types.includes(field_type)) {
            return field_type;
        }
    }

    return 'UNKNOWN';  // If no matching field type is found
}

function combineAssetData(assetData, pmaxAssets, assetPerformance) {
    let finalAssetData = {};
    assetData.forEach(row => {
        let assetId = row['asset.resource_name'];
        if (pmaxAssets[assetId]) {
            let fieldTypes = Array.from(pmaxAssets[assetId].fieldTypes);
            let primaryFieldType = fieldTypes.length === 1 ? fieldTypes[0] : select_highest_priority_field_type(fieldTypes);

            finalAssetData[assetId] = {
                name: row['asset.type'] === 'YOUTUBE_VIDEO' ? row['asset.youtube_video_asset.youtube_video_title'] : row['asset.name'],
                type: row['asset.type'] === 'YOUTUBE_VIDEO' ? 'VIDEO' : row['asset.type'],
                content: getAssetContent(row),
                performance: {},
                fieldTypes: fieldTypes,
                primaryFieldType: primaryFieldType,
                '8020': '' // Initialize 8020 field
            };

            // Preserve all campaign data and calculate 'All Campaigns'
            if (assetPerformance[assetId]) {
                for (let date in assetPerformance[assetId]) {
                    finalAssetData[assetId].performance[date] = {
                        ...assetPerformance[assetId][date],
                        'All Campaigns': {
                            impressions: 0,
                            clicks: 0,
                            cost: 0,
                            conversions: 0,
                            conversionValue: 0,
                            videoViews: 0
                        }
                    };

                    for (let campaign in assetPerformance[assetId][date]) {
                        let metrics = assetPerformance[assetId][date][campaign];
                        for (let metric in finalAssetData[assetId].performance[date]['All Campaigns']) {
                            finalAssetData[assetId].performance[date]['All Campaigns'][metric] += metrics[metric];
                        }
                    }
                }
            }

            // Log data for target asset
            if (assetId === TARGET_ID) {
                Logger.log(`Combined Asset Data: ID=${assetId}, Name=${finalAssetData[assetId].name}, Type=${row['asset.type']}`);
                Logger.log(`Field Types: ${finalAssetData[assetId].fieldTypes.join(', ')}`);
                Logger.log(`Primary Field Type: ${finalAssetData[assetId].primaryFieldType}`);
            }
        }
    });

    return calculate8020Assets(finalAssetData);
}

function calculate8020Assets(finalAssetData) {
    let assetsByType = {};
    let totalImpressionsByType = {};

    // Group assets by type and calculate total impressions
    for (let assetId in finalAssetData) {
        let asset = finalAssetData[assetId];
        asset.fieldTypes.forEach(fieldType => {
            if (!assetsByType[fieldType]) {
                assetsByType[fieldType] = [];
                totalImpressionsByType[fieldType] = 0;
            }

            let totalImpressions = Object.values(asset.performance).reduce((sum, datePerf) => {
                return sum + (datePerf['All Campaigns']?.impressions || 0);
            }, 0);

            assetsByType[fieldType].push({ id: assetId, impressions: totalImpressions });
            totalImpressionsByType[fieldType] += totalImpressions;
        });
    }

    // Sort assets by impressions and mark 8020
    for (let fieldType in assetsByType) {
        assetsByType[fieldType].sort((a, b) => b.impressions - a.impressions);

        let cumulativeImpressions = 0;
        let threshold = totalImpressionsByType[fieldType] * PERCENTAGE;

        assetsByType[fieldType].forEach(asset => {
            if (cumulativeImpressions < threshold) {
                finalAssetData[asset.id]['8020'] = 'x';
                cumulativeImpressions += asset.impressions;
            }
        });
    }

    return finalAssetData;
}

function addNamedRanges(ss, tabName, periodName) {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) {
        Logger.log(`Sheet ${tabName} not found. Skipping.`);
        return;
    }

    const timePeriod = periodName.toLowerCase();

    // Create or update named ranges using full column references
    updateNamedRange(sheet, `allDates${timePeriod}`, 'A:A');
    updateNamedRange(sheet, `name${timePeriod}`, 'B:B');
    updateNamedRange(sheet, `type${timePeriod}`, 'C:C');
    updateNamedRange(sheet, `content${timePeriod}`, 'D:D');
    updateNamedRange(sheet, `campaign${timePeriod}`, 'E:E');
    updateNamedRange(sheet, `impr${timePeriod}`, 'F:F');
    updateNamedRange(sheet, `clicks${timePeriod}`, 'G:G');
    updateNamedRange(sheet, `cost${timePeriod}`, 'H:H');
    updateNamedRange(sheet, `conv${timePeriod}`, 'I:I');
    updateNamedRange(sheet, `sub${timePeriod}`, 'L:L');
    updateNamedRange(sheet, `assetID${timePeriod}`, 'N:N');

    Logger.log(`Named ranges updated for ${tabName}`);
}

function updateNamedRange(sheet, rangeName, rangeA1Notation) {
    const ss = sheet.getParent();
    const range = sheet.getRange(rangeA1Notation);
    
    try {
        // Try to update the existing named range
        ss.setNamedRange(rangeName, range);
    } catch (e) {
        // If the named range doesn't exist, create a new one
        if (e.message.includes("Named range does not exist")) {
            ss.setNamedRange(rangeName, range);
            Logger.log(`Created new named range: ${rangeName}`);
        } else {
            // If there's a different error, log it
            Logger.log(`Error updating named range ${rangeName}: ${e.message}`);
        }
    }
}

function setInitialCampValues(ss, period, tabNames) {
    const timePeriod = period.name.toLowerCase();
    const listRangeName = `listCamp${timePeriod}`;
    const initialRangeName = `initialCamp${timePeriod}`;
    const chooseSubRangeName = `chooseSub${timePeriod}`;

    // Set initialCamp value
    const listRange = ss.getRangeByName(listRangeName);
    if (!listRange) {
        Logger.log(`Named range ${listRangeName} not found. Skipping initialCamp.`);
    } else {
        const firstValue = listRange.getValues()[0][0];
        setNamedRangeValue(ss, initialRangeName, firstValue, period, tabNames);
    }

    // Set chooseSub to HEADLINE
    setNamedRangeValue(ss, chooseSubRangeName, "HEADLINE", period, tabNames);

    Logger.log(`Finished setting values for ${timePeriod}`);
}

function setNamedRangeValue(ss, rangeName, value, period, tabNames) {
    let range = ss.getRangeByName(rangeName);
    if (!range) {
        Logger.log(`Named range ${rangeName} not found. Creating it.`);
        const sheet = ss.getSheetByName(tabNames[period.name]);
        if (!sheet) {
            Logger.log(`Sheet ${tabNames[period.name]} not found. Skipping.`);
            return;
        }
        // Create the named range in the next available cell in column A
        const nextRow = sheet.getLastRow() + 1;
        const newRange = sheet.getRange(`A${nextRow}`);
        ss.setNamedRange(rangeName, newRange);
        range = newRange;
    }

    // Set the value
    range.setValue(value);
    Logger.log(`Set ${rangeName} to value: ${value}`);
}

function writeDataToSheet(ss, finalAssetData, tabName) {
    Logger.log(`Writing data to ${tabName}`);
    let sheet = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
    sheet.clearContents();

    // Define headers
    let headers = ['Date', 'Name', 'Type', 'Content', 'Campaign', 'Impr', 'Clicks', 'Cost', 'Conv', 'Value', 'Views', 'Sub Type', '8020', 'Asset ID', 'Possible Field Types'];
    let data = [headers];

    // Process finalAssetData to build all the rows in one go
    for (let assetId in finalAssetData) {
        let asset = finalAssetData[assetId];
        for (let date in asset.performance) {
            // Prepare name and content fields
            let name = asset.type === 'TEXT' ? asset.content : asset.name;
            let content = asset.type === 'TEXT' ? '' :
                (asset.type === 'YOUTUBE_VIDEO' ? `https://www.youtube.com/watch?v=${asset.content}` : asset.content);

            // Add rows for individual campaigns
            for (let campaign in asset.performance[date]) {
                if (campaign !== 'All Campaigns') {
                    let metrics = asset.performance[date][campaign];
                    data.push([
                        date,
                        name,
                        asset.type,
                        content,
                        campaign,
                        metrics.impressions,
                        metrics.clicks,
                        metrics.cost,
                        metrics.conversions,
                        metrics.conversionValue,
                        metrics.videoViews,
                        asset.primaryFieldType,
                        '', // Empty 8020 column for individual campaigns
                        assetId,
                        asset.fieldTypes.join(', ')
                    ]);
                }
            }

            // Add the 'All Campaigns' summary row
            let allCampaignsMetrics = asset.performance[date]['All Campaigns'];
            data.push([
                date,
                name,
                asset.type,
                content,
                'All Campaigns',
                allCampaignsMetrics.impressions,
                allCampaignsMetrics.clicks,
                allCampaignsMetrics.cost,
                allCampaignsMetrics.conversions,
                allCampaignsMetrics.conversionValue,
                allCampaignsMetrics.videoViews,
                asset.primaryFieldType,
                asset['8020'] || '', // 8020 value for 'All Campaigns' row
                assetId,
                asset.fieldTypes.join(', ')
            ]);
        }
    }

    // Write data to the sheet in one go
    if (data.length > 1) { // Ensure we have data to write (not just headers)
        sheet.getRange(1, 1, data.length, headers.length).setValues(data);
    } else {
        Logger.log(`No data to write to sheet ${tabName}.`);
    }
}

// thanks for using this script

// ps you're awesome
