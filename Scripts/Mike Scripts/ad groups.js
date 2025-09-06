function main() {
    const report = AdsApp.report(`
      SELECT 
        campaign.name,
        ad_group.name,
        metrics.clicks,
        metrics.impressions,
        metrics.cost_micros,
        metrics.conversions
      FROM ad_group 
      WHERE segments.date DURING LAST_30_DAYS
      ORDER BY metrics.cost_micros DESC
      LIMIT 10
    `);
    
    const rows = report.rows();
    
    while (rows.hasNext()) {
      const row = rows.next();
      const cost = (row['metrics.cost_micros'] / 1000000).toFixed(2);
      
      console.log(`${row['campaign.name']} > ${row['ad_group.name']}`);
      console.log(`Clicks: ${row['metrics.clicks']} | Cost: $${cost} | Conversions: ${row['metrics.conversions']}`);
      console.log('---');
    }
  }