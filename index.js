const runQuery = (query) => {
  const projectId = "";

  const request = {
    query,
    useLegacySql: false,
  };
  Logger.log(request);
  let queryResults = BigQuery.Jobs.query(request, projectId);
  const { jobId } = queryResults.jobReference;

  let sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *= 2;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }

  let { rows } = queryResults;
  while (queryResults.pageToken) {
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
      pageToken: queryResults.pageToken,
    });
    rows = rows.concat(queryResults.rows);
  }

  if (rows) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    const range = sheet.getRange("A:ZZ");
    range.clearContent();

    const headers = queryResults.schema.fields.map((field) => field.name);
    sheet.appendRow(headers);

    const data = new Array(rows.length);
    for (let i = 0; i < rows.length; i++) {
      const cols = rows[i].f;
      data[i] = new Array(cols.length);
      for (let j = 0; j < cols.length; j++) {
        data[i][j] = cols[j].v;
      }
    }
    sheet.getRange(2, 1, rows.length, headers.length).setValues(data);

    Logger.log("Results spreadsheet created: %s", spreadsheet.getUrl());
  } else {
    Logger.log("No rows returned.");
  }
};

const SalesOrderLines = () =>
  runQuery(`
      SELECT *
      FROM NetSuite.vn_SalesOrderLines
      WHERE DATE(TRANDATE) >= CURRENT_DATE();`);

const SalesLines = () =>
  runQuery(`
      SELECT *
      FROM NetSuite.vn_SalesLines
      WHERE DATE(TRANDATE) >= CURRENT_DATE();`);

const COGSLines = () =>
  runQuery(`
      SELECT *
      FROM NetSuite.vn_COGSLines
      WHERE DATE(TRANDATE) >= CURRENT_DATE();`);

const onOpen = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const menuEntries = [
    { name: "SalesOrderLines", functionName: "SalesOrderLines" },
    { name: "SalesLines", functionName: "SalesLines" },
    { name: "COGSLines", functionName: "COGSLines" },
  ];
  sheet.addMenu("Vua Nem BigQuery", menuEntries);
};
