
// Initialize MSAL Configuration
const msalConfig = {
    auth: {
        clientId: "<your-client-id-goes-here>", // Replace with your Azure AD App's Client ID
        authority: "https://login.microsoftonline.com/<your-tenant-id-goes-here>", // Replace with Tenant ID
        redirectUri: "http://localhost:8000", // Replace with your Redirect URI
    },
    cache: {
        cacheLocation: "localStorage", // Stores tokens in localStorage
        storeAuthStateInCookie: false, // Set true for older browsers
    },
};


// Create MSAL instance
let msalInstance;
try {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    console.log("MSAL Instance initialized successfully.");
} catch (error) {
    console.error("Error initializing MSAL instance:", error);
}

// Login function
async function login() {
    try {
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["User.ReadWrite.All", "Directory.ReadWrite.All", "Mail.Send"],
        });
        msalInstance.setActiveAccount(loginResponse.account);
        alert("Login successful.");
    } catch (error) {
        console.error("Login error:", error);
        alert("Login failed.");
    }
}

// Logout function
function logout() {
    msalInstance.logoutPopup().then(() => alert("Logout successful."));
}

// Retrieve disabled users
async function retrieveDisabledUsers() {
    try {
        const response = await callGraphApi(`/users?$filter=accountEnabled eq false&$select=userPrincipalName,givenName,surname,department,jobTitle`);
        populateTable(response.value);
    } catch (error) {
        console.error("Error retrieving disabled users:", error);
        alert("Failed to retrieve disabled users.");
    }
}


async function filterUsersByLicense() {
    const licenseStatus = document.getElementById("licenseStatusFilter").value;

    if (!licenseStatus) {
        alert("Please select a license status.");
        return;
    }

    try {
        // Fetch all disabled users
        const response = await callGraphApi(`/users?$filter=accountEnabled eq false&$select=userPrincipalName,givenName,surname,department,jobTitle,assignedLicenses`);

        // Filter users based on license status client-side
        const filteredUsers = response.value.filter(user => {
            const hasLicense = user.assignedLicenses && user.assignedLicenses.length > 0;
            return licenseStatus === "licensed" ? hasLicense : !hasLicense;
        });

        if (filteredUsers.length > 0) {
            populateTable(filteredUsers);
        } else {
            alert("No disabled users found for the selected license status.");
            clearTable();
        }
    } catch (error) {
        console.error("Error filtering users by license status:", error);
        alert("Failed to filter users by license status.");
    }
}

// Clear table data
function clearTable() {
    document.getElementById("outputHeader").innerHTML = "";
    document.getElementById("outputBody").innerHTML = "";
}



// Populate table
function populateTable(users) {
    const outputHeader = document.getElementById("outputHeader");
    const outputBody = document.getElementById("outputBody");

    outputHeader.innerHTML = `
        <th>UserPrincipalName</th>
        <th>Signin Status</th>
        <th>First Name</th>
        <th>Last Name</th>
        <th>Department</th>
        <th>Job Title</th>
    `;

    outputBody.innerHTML = users.map(user => `
        <tr>
            <td>${user.userPrincipalName || "N/A"}</td>
            <td>Disabled</td>
            <td>${user.givenName || "N/A"}</td>
            <td>${user.surname || "N/A"}</td>
            <td>${user.department || "N/A"}</td>
            <td>${user.jobTitle || "N/A"}</td>
        </tr>
    `).join("");
}

// Search users


// Search Users by Multiple Properties (Disabled Users Only)
async function searchUsers() {
    const query = document.getElementById("searchInput").value.trim();

    if (!query) {
        alert("Please enter a search query.");
        return;
    }

    // Filter query to search for disabled users only
    const filterQuery = `
        accountEnabled eq false and (
            startswith(userPrincipalName,'${query}') or
            startswith(givenName,'${query}') or
            startswith(surname,'${query}') or
            startswith(displayName,'${query}') or
            startswith(mail,'${query}')
        )
    `.trim();

    try {
        const encodedFilter = encodeURIComponent(filterQuery);
        const response = await callGraphApi(`/users?$filter=${encodedFilter}&$select=userPrincipalName,givenName,surname,department,jobTitle`);

        if (response.value && response.value.length > 0) {
            populateTable(response.value);
        } else {
            alert("No disabled users found for the given search query.");
            clearTable();
        }
    } catch (error) {
        console.error("Error searching users:", error);
        alert("An error occurred while searching for users. Please try again.");
    }
}

// Clear table data
function clearTable() {
    document.getElementById("outputHeader").innerHTML = "";
    document.getElementById("outputBody").innerHTML = "";
}









// Send report as mail
// Send Report as Mail
async function sendReportAsMail() {
    const recipientEmail = document.getElementById("recipientEmail").value;

    if (!recipientEmail) {
        alert("Please enter a valid recipient email.");
        return;
    }

    // Extract data from the table
    const tableHeaders = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const tableRows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (tableRows.length === 0) {
        alert("No data to send. Please retrieve and display user details first.");
        return;
    }

    // Format the email body as an HTML table
    const emailTable = `
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <thead>
                <tr>${tableHeaders.map(header => `<th>${header}</th>`).join("")}</tr>
            </thead>
            <tbody>
                ${tableRows
                    .map(
                        row => `<tr>${row.map(cell => `<td>${cell}</td>`).join("")}</tr>`
                    )
                    .join("")}
            </tbody>
        </table>
    `;

    // Email content
    const email = {
        message: {
            subject: "User Report from M365 User Management Tool",
            body: {
                contentType: "HTML",
                content: `
                    <p>Dear Administrator,</p>
                    <p>Please find below the user report generated by the M365 User Management Tool:</p>
                    ${emailTable}
                    <p>Regards,<br>M365 User Management Team</p>
                `
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: recipientEmail
                    }
                }
            ]
        }
    };

    try {
        const response = await callGraphApi("/me/sendMail", "POST", email);
        alert("Report sent successfully!");
        console.log("Mail Response:", response);
    } catch (error) {
        console.error("Error sending report:", error);
        alert("Failed to send the report. Please try again.");
    }
}
// Download report as CSV
function downloadReportAsCSV() {
    const headers = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (rows.length === 0) {
        alert("No data to download.");
        return;
    }

    const csvContent = [headers.join(","), ...rows.map(r => r.join(","))].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "Disabled_Users_Report.csv";
    a.click();
    URL.revokeObjectURL(url);
}

// Reset screen
function resetScreen() {
    document.getElementById("searchInput").value = "";
    document.getElementById("recipientEmail").value = "";
    document.getElementById("licenseStatusFilter").value = "";
    document.getElementById("outputHeader").innerHTML = "";
    document.getElementById("outputBody").innerHTML = "";
    alert("Screen has been reset.");
}

// Call Graph API

async function callGraphApi(endpoint, method = "GET", body = null) {
    const account = msalInstance.getActiveAccount();
    if (!account) {
        throw new Error("No active account. Please login first.");
    }

    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["User.ReadWrite.All", "Directory.ReadWrite.All", "Mail.Send"],
            account: account,
        });

        const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
            method,
            headers: {
                Authorization: `Bearer ${tokenResponse.accessToken}`,
                "Content-Type": "application/json",
            },
            body: body ? JSON.stringify(body) : null,
        });

        if (response.ok) {
            const contentType = response.headers.get("content-type");
            if (contentType && contentType.includes("application/json")) {
                return await response.json(); // Parse JSON responses
            }
            return {}; // Return empty object for non-JSON responses (e.g., 204 No Content)
        } else {
            const errorText = await response.text(); // Read error as text
            console.error("Graph API error response:", errorText);
            throw new Error(`Graph API call failed: ${response.status} ${response.statusText}`);
        }
    } catch (error) {
        console.error("Error in callGraphApi:", error.message);
        throw error;
    }
}

