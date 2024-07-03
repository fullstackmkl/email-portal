let mailtoLinks = []; // Array to store mailto links for all emails
let jobData = []; // Array to store the job data
let userData = []; // Array to store the user data
let currentUser = null; // Store the current user information

// Example file paths for the Excel files
const emailJobsFilePath = 'Z:\\OpReef\\email jobs.xlsx';
const userCredentialsFilePath = 'Z:\\OpReef\\credentials.xlsx';

// Function to login the user
function login() {
  const username = document.getElementById('username').value;
  const password = document.getElementById('password').value;
  const loginError = document.getElementById('loginError');

  // Load the user credentials from the Excel file
  loadExcelFile(userCredentialsFilePath, (data) => {
    userData = data.slice(1); // Exclude the headers

    // Find the user in the user data
    currentUser = userData.find(user => user[0] === username && user[1] === password);

    if (currentUser) {
      loginError.style.display = 'none';
      document.getElementById('loginSection').style.display = 'none';
      document.getElementById('jobSection').style.display = 'block';
      loadEmailJobs();
    } else {
      loginError.textContent = 'Invalid username or password.';
      loginError.style.display = 'block';
    }
  });
}

// Function to load the email jobs from the Excel file
function loadEmailJobs() {
  loadExcelFile(emailJobsFilePath, (data) => {
    jobData = data.slice(1); // Exclude the headers
    displayJobTable();
  });
}

// Function to load an Excel file and parse it into JSON
function loadExcelFile(filePath, callback) {
  const xhr = new XMLHttpRequest();
  xhr.open('GET', filePath, true);
  xhr.responseType = 'arraybuffer';
  xhr.onload = function (e) {
    const data = new Uint8Array(xhr.response);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    callback(jsonData);
  };
  xhr.send();
}

// Function to display the job table
function displayJobTable() {
  const jobTableContainer = document.getElementById('jobTableContainer');
  jobTableContainer.innerHTML = '';

  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');

  // Create table headers
  const headers = ['Name', 'Description', 'Status', 'Client Name', 'Time To Complete', 'Action'];
  const headerRow = document.createElement('tr');
  headers.forEach(header => {
    const th = document.createElement('th');
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  // Create table rows
  jobData.forEach((job, index) => {
    const tr = document.createElement('tr');
    job.forEach((cell, cellIndex) => {
      const td = document.createElement('td');
      td.textContent = cell;
      if (cellIndex === 2) {
        // Status column
        const select = document.createElement('select');
        select.innerHTML = '<option value="Enabled">Enabled</option><option value="Disabled">Disabled</option>';
        select.value = cell;
        select.onchange = () => updateJobStatus(index, select.value);
        td.innerHTML = '';
        td.appendChild(select);
      } else {
        td.contentEditable = true; // Make other cells editable
      }
      tr.appendChild(td);
    });

    // Add action button for file upload
    const td = document.createElement('td');
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx, .xls';
    fileInput.onchange = (e) => handleFileUpload(e, index);
    td.appendChild(fileInput);
    tr.appendChild(td);

    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  jobTableContainer.appendChild(table);
}

// Function to update the job status
function updateJobStatus(index, status) {
  jobData[index][2]
 = status; // Update the status in the job data
  saveExcelFile(emailJobsFilePath, jobData, () => {
    console.log('Job status updated');
  });
}

// Function to handle file upload
function handleFileUpload(event, index) {
  const file = event.target.files[0];
  if (!file.name.match(/\.(xlsx|xls)$/)) {
    alert('Invalid file type. Please upload an Excel file.');
    return;
  }
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    jobData[index].push(jsonData); // Store the uploaded data in the job data
  };
  reader.readAsArrayBuffer(file);
}

// Function to save the Excel file with updated data
function saveExcelFile(filePath, data, callback) {
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  const wopts = { bookType: 'xlsx', type: 'array' };
  const wbout = XLSX.write(workbook, wopts);
  const blob = new Blob([wbout], { type: 'application/octet-stream' });

  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filePath.split('/').pop();
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  callback();
}

// Function to run email jobs
function runEmailJobs() {
  const emailDraftsContainer = document.getElementById('emailDrafts');
  emailDraftsContainer.innerHTML = '';
  mailtoLinks = [];

  jobData.forEach(job => {
    if (job[2] === 'Enabled' && job[5]) { // Check if the job is enabled and has uploaded data
      const emailBody = generateEmailBody(job[5][0], job[5].slice(1), job[0]); // Generate email body using job name as notification type
      const emailColumnIndex = job[5][0].indexOf('Email');
      const emails = [...new Set(job[5].slice(1).map(row => row[emailColumnIndex]))]; // Get unique emails
      emails.forEach(email => {
        const emailLink = `mailto:${email}?subject=${getSubject(job[0])}&body=${encodeURIComponent(emailBody)}`;
        mailtoLinks.push(emailLink);
        appendEmailDraft(emailDraftsContainer, email, emailBody, emailLink);
      });
    }
  });

  document.getElementById('emailSection').style.display = 'block';
}

// Function to generate the email body based on the job data
function generateEmailBody(headers, items, notificationType) {
  let body = `Dear Manager,\n\n`;

  switch (notificationType) {
    case 'Timesheets Pending Approval':
      body += 'This is a reminder to submit your timesheets. Please review and approve the following timesheets:\n\n';
      break;
    case 'Contracts Expiring':
      body += 'This is a notification for expiring contracts. Please review and notify if an extension is needed:\n\n';
      break;
    case 'Audit Reminders':
      body += 'This is an audit reminder. Please review the following timesheets for compliance:\n\n';
      break;
    default:
      body += 'Please review the following:\n\n';
  }

  body += headers.join('\t') + '\n';
  body += headers.map(() => '---').join('\t') + '\n';
  items.forEach(item => {
    body += item.join('\t') + '\n';
  });

  body += `\nThank you,\nYour Automated System`;
  return body;
}

// Function to get the subject of the email based on the job name
function getSubject(jobName) {
  switch (jobName) {
    case 'Timesheets Pending Approval':
      return 'Timesheet Approval Needed';
    case 'Contracts Expiring':
      return 'Contract Expiration Notification';
    case 'Audit Reminders':
      return 'Audit Reminder';
    default:
      return 'Notification';
  }
}

// Function to append an email draft to the container
function appendEmailDraft(container, email, body, mailtoLink) {
  const emailDraft = document.createElement('div');
  emailDraft.classList.add('email-draft');
  emailDraft.innerHTML = `
    <p><strong>To:</strong> ${email}</p>
    <p><strong>Subject:</strong> ${getSubject(document.getElementById('notificationType').value)}</p>
    <p><strong>Body:</strong></p>
    <pre>${body}</pre>
    <a href="${mailtoLink}" target="_blank">Open in Outlook</a>
  `;
  container.appendChild(emailDraft);
}

// Function to send all emails
function sendAllEmails() {
  mailtoLinks.forEach(link => {
    window.open(link, '_blank');
  });
}
