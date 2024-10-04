require("dotenv").config();
const axios = require("axios");
const ExcelJS = require("exceljs");
const {
  startOfMonth,
  endOfMonth,
  format,
  addMonths,
  differenceInMonths,
  parseISO,
  parse,
  isValid,
} = require("date-fns");

const jiraAccessToken = process.env.JIRA_ACCESS_TOKEN;
const assigneeNames = process.env.JIRA_TEAM_MEMBERS.split(",");
const jiraUrl = process.env.JIRA_URL;
const jiraProjectName = process.env.JIRA_PROJECT_NAME;
const jiraStatus = process.env.JIRA_STATUS;
const jiraStatusCategory = process.env.JIRA_STATUS_CATEGORY;
const jiraOrderBy = process.env.JIRA_ORDER_BY;
const jiraOrderDirection = process.env.JIRA_ORDER_DIRECTION;
const jiraTeamName = process.env.JIRA_TEAM_NAME;
const jiraStartMonth = process.env.JIRA_START_MONTH;
const jiraEndMonth = process.env.JIRA_END_MONTH;
const overrideExcelDateFormat = process.env.EXCEL_DATE_FORMAT;

const currentDate = new Date();

const parseDateOrMonth = (dateStr) => {
  let date = parseISO(dateStr);
  if (!isValid(date)) {
    date = parse(dateStr, "yyyy-MM", new Date());
  }
  return date;
};

const monthStart = jiraStartMonth
  ? startOfMonth(parseDateOrMonth(jiraStartMonth))
  : startOfMonth(currentDate);
const monthEnd = jiraEndMonth
  ? endOfMonth(parseDateOrMonth(jiraEndMonth))
  : endOfMonth(currentDate);

const dateFormat = "yyyy-MM-dd";
const excelDateFormat = overrideExcelDateFormat + " HH:mm";

console.log(
  "Querying [" + assigneeNames.length + "] employees for the following dates:"
);
console.log("Start date", format(monthStart, dateFormat));
console.log("End date", format(monthEnd, dateFormat));

const workbook = new ExcelJS.Workbook();

// const fetchIssuesForAssignee = async (
//   assignee,
//   startDate,
//   endDate,
//   nextMonthFirstDay
// ) => {
//   const config = {
//     method: "get",
//     maxBodyLength: Infinity,
//     url: `${jiraUrl}/rest/api/2/search?jql=assignee was in (${assignee}) during ("${format(
//       startDate,
//       dateFormat
//     )}", "${format(endDate, dateFormat)}") AND status was not in (${jiraStatus}) before "${format(
//       startDate,
//       dateFormat
//     )}" AND status was in (${jiraStatus}) before "${format(
//       nextMonthFirstDay,
//       dateFormat
//     )}" AND statusCategory = ${jiraStatusCategory} ORDER BY ${jiraOrderBy} ${jiraOrderDirection}`,
//     headers: {
//       Authorization: `Bearer ${jiraAccessToken}`,
//     },
//   };


const fetchIssuesForAssignee = async (
  assignee,
  startDate,
  endDate,
  nextMonthFirstDay
) => {
  const config = {
    method: "get",
    maxBodyLength: Infinity,
    url: `${jiraUrl}/rest/api/2/search?jql=assignee was in (${assignee}) during ("${format(
      startDate,
      dateFormat
    )}", "${format(endDate, dateFormat)}")`,
    headers: {
      Authorization: `Bearer ${jiraAccessToken}`,
    },
  };



  console.log(config.url);

  try {
    const response = await axios.request(config);
    return response.data.issues;
  } catch (error) {
    console.error(error.message);
    return [];
  }
};

const applyConditionalFill = (value) => {
  if (/done/i.test(value)) {
    return { type: "pattern", pattern: "solid", fgColor: { argb: "ffccffce" } }; // Light green color
  } else if (/won't do/i.test(value) || /closed/i.test(value)) {
    return { type: "pattern", pattern: "solid", fgColor: { argb: "FFD3D3D3" } }; // Light gray color
  }
  return null;
};

const applyHeaderStyles = (worksheet, headers) => {
  const borderStyle = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };

  const customStyles = {
    month: { fgColor: { argb: "ffffff00" } }, // Yellow
    year: { fgColor: { argb: "fff8e5d5" } }, // Light pink
    number_of_task: { fgColor: { argb: "ffdbedf4" } }, // Light blue
  };

  const headerRow = worksheet.getRow(1);

  headers.forEach((header, index) => {
    const cell = headerRow.getCell(index + 1);
    cell.font = { bold: true };
    cell.border = borderStyle;
    if (customStyles[header.key]) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: customStyles[header.key].fgColor,
      };
    }
  });
  headerRow.height = 30;
  headerRow.alignment = { vertical: "middle" };
  headerRow.commit();
};

const createWorksheetForMonth = async (issuesByAssignee, month, year) => {
  const worksheet = workbook.addWorksheet(`${month}-${year}`);

  const headers = [
    { header: "Project", key: "project", width: 20 },
    { header: "Assignee", key: "assignee", width: 20 },
    { header: "Issue Key", key: "key", width: 20 },
    { header: "Issue ID", key: "issue_id", width: 20 },
    { header: "Parent ID", key: "parent_id", width: 20 },
    { header: "Summary", key: "summary", width: 32 },
    { header: "Status", key: "status", width: 15 },
    { header: "Created Date", key: "created", width: 20 },
    { header: "Updated Date", key: "updated", width: 20 },
    { header: "Month-GTV", key: "month", width: 20 },
    { header: "Year-GTV", key: "year", width: 20 },
    { header: "Number of Task", key: "number_of_task", width: 20 },
    { header: "Issue Type", key: "issuetype", width: 20 },
    { header: "Resolution", key: "resolution", width: 20 },
  ];

  worksheet.columns = headers;

  applyHeaderStyles(worksheet, headers);

  worksheet.autoFilter = { from: "A1", to: "N1" };

  issuesByAssignee.forEach((assigneeIssues, index) => {
    assigneeIssues.issues.forEach((issue) => {
      const row = worksheet.addRow({
        project: jiraProjectName,
        assignee: assigneeIssues.assignee,
        key: issue.key,
        issue_id: issue.id,
        parent_id: issue.fields.parent ? issue.fields.parent.id : "",
        summary: issue.fields.summary,
        status: issue.fields.status.name,
        created: format(new Date(issue.fields.created), excelDateFormat),
        updated: format(new Date(issue.fields.resolutiondate), excelDateFormat),
        month: format(new Date(issue.fields.resolutiondate), "MMMM"),
        year: format(new Date(issue.fields.resolutiondate), "yyyy"),
        number_of_task: 1,
        issuetype: issue.fields.issuetype.name,
        resolution: issue.fields.resolution?.name || "Unresolved",
      });

      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });

      // Apply conditional fill to the "Status" cell
      const statusCell = row.getCell("status");
      const statusFillStyle = applyConditionalFill(issue.fields.status.name);
      if (statusFillStyle) {
        statusCell.fill = statusFillStyle;
      }

      // Apply conditional fill to the "Resolution" cell
      const resolutionCell = row.getCell("resolution");
      const fillStyle = applyConditionalFill(
        issue.fields.resolution?.name || ""
      );
      if (fillStyle) {
        resolutionCell.fill = fillStyle;
      }
    });

    if (index < issuesByAssignee.length - 1) {
      const emptyRow = worksheet.addRow({
        project: "",
        assignee: "",
        key: "",
        issue_id: "",
        parent_id: "",
        summary: "",
        status: "",
        created: "",
        updated: "",
        month: "",
        year: "",
        issuetype: "",
        resolution: "",
      });
      emptyRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "ffff00" },
      };

      emptyRow.commit();
    }
  });

  worksheet.font = { name: "Calibri", size: 11 };
};

const processAssignees = async () => {
  const totalMonths = differenceInMonths(monthEnd, monthStart) + 1;

  for (let i = 0; i < totalMonths; i++) {
    const startDate = addMonths(monthStart, i);
    const endDate = endOfMonth(startDate);
    const nextMonthFirstDay = getNextMonthFirstDay(endDate);

    const issuesByAssignee = await Promise.all(
      assigneeNames.map(async (assignee) => {
        console.log(`Retrieving information for dates:
          Assignee: ${assignee}
          Start Date: ${format(startDate, "yyyy-MM-dd HH:mm")}
          End Date: ${format(endDate, "yyyy-MM-dd HH:mm")}
          Next Month First Day: ${format(nextMonthFirstDay, "yyyy-MM-dd HH:mm")}
        `);

        const issues = await fetchIssuesForAssignee(
          assignee,
          startDate,
          endDate,
          nextMonthFirstDay
        );
        return { assignee, issues };
      })
    );

    // Filter out empty issues
    const nonEmptyIssuesByAssignee = issuesByAssignee.filter(
      (assigneeIssues) => assigneeIssues.issues.length > 0
    );

    if (nonEmptyIssuesByAssignee.length > 0) {
      await createWorksheetForMonth(
        nonEmptyIssuesByAssignee,
        format(startDate, "MMMM"),
        format(startDate, "yyyy")
      );
    }
  }

  await workbook.xlsx.writeFile(
    `${jiraTeamName}_Report_${format(monthStart, dateFormat)}_${format(
      monthEnd,
      dateFormat
    )}.xlsx`
  );
  console.log("Data exported to Excel file successfully.");
};

function getNextMonthFirstDay(date) {
  let newDate = new Date(date);

  // Reset first to prevent adding 2 months if the day does not exist in the following month
  newDate.setDate(1); // Set to the first day of the month
  newDate.setMonth(newDate.getMonth() + 1); // Add one month

  return newDate;
}

// main
processAssignees();
