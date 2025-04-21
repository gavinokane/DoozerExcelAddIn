/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

let data_dictionary = {}; // Global data dictionary to store range info

function stubBackendCall(instance_id) {
  // Simulate a backend call to check the status of an instance_id
  const status = ["ready"];
  
  CheckStatusOfInstance(instance_id);
  
  return status;
}


function pollStatus() {
  Excel.run(async (context) => {

    let allDone = true;

    const progressBar = document.getElementById("progress-bar");
    if (!progressBar) throw new Error("Progress bar element not found.");
    progressBar.value = progressBar.value + 10;
    
    progressBar.className = ""; // Clear existing classes if any
    progressBar.classList.add("progress-waiting"); // Add a class for styling

    for (const [key, value] of Object.entries(data_dictionary)) {
      // Log the current status and instance_id
      AddtoLog(`Range: ${key},  Instance ID: ${value.instance_id}, Status: ${value.status},`);
      
      if (value.status === "pending") {
        allDone = false;

        // Await the Promise returned by CheckStatusOfInstance
        const isUpdated = await CheckStatusOfInstance(value.instance_id);

        if (isUpdated === true) {
          AddtoLog(`Wrote answer to cell to the right of range ${key}`);
          data_dictionary[key].status = "done";
        }
      } else if (value.status === "ready") {
        allDone = false;
        AddtoLog(`Range ${key} is marked as ready. Retrying in the next poll.`);
      }
    }

    // If all entries are done, set progress bar to green and 100%
    if (allDone) {
      const progressBar = document.getElementById("progress-bar");
      if (!progressBar) throw new Error("Progress bar element not found.");
      progressBar.value = 100;// When process is complete
      progressBar.classList.remove("progress-waiting");
    }
  }).catch((error) => {
    console.error("Error in pollStatus:", error);
    AddtoLog(`Error in pollStatus: ${error.message}`);
  });
}

Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

 
  // Login button handler
  document.getElementById("login-button").onclick = handleLogin;

  // Logout button handler
  document.getElementById("logout-button").onclick = handleLogout;



  // Fetch worker list on login
  const subscriptionKey = localStorage.getItem("subscriptionKey");
  const apiKey = localStorage.getItem("apiKey");
  if (subscriptionKey && apiKey) {
    fetchWorkerList(subscriptionKey, apiKey);
//document.getElementById("worker-section").innerHTML = ""; // Clear existing content
fetchWorkerList(subscriptionKey, apiKey); // Refresh worker list
    document.getElementById("login-frame").style.display = "none";
} else {
    document.getElementById("login-frame").style.display = "block";
  }

  document.getElementById("run").onclick = SubmitToAgent;

  const progressBar = document.getElementById("progress-bar");
  if (progressBar) {
    progressBar.value = 100;
    progressBar.style.backgroundColor = "green"; // Initialize as green and 100%
  }

  // Attach handlers for the new buttons
  


  // Attach handler for the "Clear Log" button
  document.getElementById("clear-log").onclick = clearLog;
  document.getElementById("clear-jobs").onclick = clearJobs;
  document.getElementById("status").onclick = status;
  

  // Log the origin URL and IP address during initialization
  const origin = window.location.origin;
  AddtoLog(`Origin URL: ${origin}`);

  async function logIPAddress() {
    try {
      const ipResponse = await fetch("https://api.ipify.org?format=json");
      if (ipResponse.ok) {
        const ipData = await ipResponse.json();
        AddtoLog(`IP Address: ${ipData.ip}`);
      } else {
        AddtoLog(`Failed to fetch IP address. Status: ${ipResponse.status}`);
      }
    } catch (error) {
      AddtoLog(`Error fetching IP address: ${error.message}`);
    }
  }

  logIPAddress();

  // Set up polling mechanism to run every 20 seconds
  setInterval(pollStatus, 5000);
});


async function updateExcelCell(sheetName, cellReference, answer) {
  try {
    await Excel.run(async (context) => {
      AddtoLog(`Attempting to write to ${sheetName}!${cellReference}`);
      
      // Get the sheet by name
      let sheet;
      try {
        sheet = context.workbook.worksheets.getItem(sheetName);
      } catch (sheetError) {
        AddtoLog(`ERROR: Sheet '${sheetName}' not found. ${sheetError.message}`);
        return;
      }
      
      // Get the cell and update it
      try {
        const cell = sheet.getRange(cellReference);
        try {
          cell.values = [[answer]];
          cell.format.autofitRows();
          await context.sync();
          AddtoLog(`Successfully wrote to ${sheetName}!${cellReference}`);
        } catch (cellError) {
          if (cellError.message.includes("Excel is in cell-editing mode")) {
            AddtoLog(`Update canceled: Excel is in cell-editing mode for ${sheetName}!${cellReference}`);
            if (data_dictionary[cellReference]) {
              data_dictionary[cellReference].status = "ready";
            }
            return;
          }
          throw cellError;
        }
      } catch (cellError) {
        AddtoLog(`ERROR: Could not update cell ${cellReference}. ${cellError.message}`);
      }
    });
  } catch (error) {
    AddtoLog(`ERROR in updateExcelCell: ${error.message}`);
    console.error(`Error updating cell ${cellReference} on sheet ${sheetName}:`, error);
  }
}


// / Helper function to convert column number to letter
function getColumnLetter(column) {
    let temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - (temp + 1)) / 26;
    }
    return letter;
}

function processAnswers(jsonString) {
  try {
    // Extract just the JSON portion from the string
    const jsonStart = jsonString.indexOf('{');
    const jsonEnd = jsonString.lastIndexOf('}') + 1;
    const jsonPortion = jsonString.substring(jsonStart, jsonEnd);
    
    // Parse the JSON
    const data = JSON.parse(jsonPortion);
    
    // Extract sheet name from the address
    const addressMatch = data.address.match(/\'([^\']+)\'/);
    const sheetName = addressMatch ? addressMatch[1] : data.address.split('!')[0];
    
    AddtoLog(`Processing answers for sheet: ${sheetName}`);
    
    // Skip the header row if it exists
    const answers = Array.isArray(data.values[0][0]) ? data.values.slice(1) : data.values;
    
    // Loop through each question/answer pair
    answers.forEach((pair, index) => {
      const question = pair[0];
      const answer = pair[1];
      
      // Get row and column information
      const row = data.rowIndex + index + 1; // Adjust based on your data structure
      const column = data.columnIndex + 2; // Column to the right of the question
      const colLetter = getColumnLetter(column);
      const cellReference = `${colLetter}${row}`;
      
      AddtoLog(`Writing answer to ${sheetName}!${cellReference}: ${answer}`);
      
      // Update the Excel cell
      updateExcelCell(sheetName, cellReference, answer);
    });
    
    return answers;
  } catch (error) {
    AddtoLog(`ERROR processing answers: ${error.message}`);
    console.error("Error processing JSON:", error);
    return null;
  }
}


function CheckStatusOfInstance(instance_id) {
  return new Promise((resolve, reject) => {
    const myHeaders = new Headers();
const subscriptionKey = localStorage.getItem("subscriptionKey");
const apiKey = localStorage.getItem("apiKey");

if (subscriptionKey && apiKey) {
    myHeaders.append("Ocp-Apim-Subscription-Key", subscriptionKey);
    myHeaders.append("API_KEY", apiKey);
} else {
    console.error("Error: Keys are not available in local storage.");
}

    const requestOptions = {
      method: "GET",
      headers: myHeaders,
      redirect: "follow"
    };

    // AddtoLog(`Checking status for instance ${instance_id}...`);

    fetch(`https://fn-doozer-py-05.azurewebsites.net/api/Workflow/Instance?instance_id=${instance_id}`, requestOptions)
      .then((response) => {
        // AddtoLog(`Response status: ${response.status}`);
        return response.text();
      })
      .then((result) => {
        // Check if result is null or the string "null"
        if (result === null || result === undefined || result === "null" || result.trim() === "") {
          resolve(false);
          return;
        }
        
        try {
          const data = JSON.parse(result);

          if (data.status !== "complete") {
            AddtoLog(`Agent instance ${instance_id} is ${data.status}`);
            resolve(false);
            return;
          }

          // Fix the typo in property name
          const dataDictionary = data.data_dictionary || data.data_dictinary;
          if (dataDictionary) {
            // AddtoLog("Data Dictionary found");
            AddtoLog(dataDictionary.final_answer);
            
            processAnswers(dataDictionary.final_answer);
            resolve(true);
          } else {
            AddtoLog("Data Dictionary not found in response");
            resolve(false);
          }
        } catch (error) {
          AddtoLog("Error parsing response: " + error);
          resolve(false);
        }
      })
      .catch((error) => {
        AddtoLog("Error in fetch: " + error);
        resolve(false);
      });
  });
}
function AddtoLog(message) {
    const logWindow = document.getElementById("log-window");
    if (logWindow) {
        const logMessage = document.createElement("p");
        logMessage.textContent = message;
        logWindow.appendChild(logMessage);
    } else {
        console.error("Log window not found.");
    }
}

// // Handler for the "AutoFill" button
// function handleAutoFill() {
   
//   const payload = {"status":"pending","instance_id":"","columnIndex":3}
  
//   data_dictionary['Sheet1!D19:E22'] = {
//     status: "pending",
//     instance_id: "85d4947d-d24c-4f1a-81bf-35ccceccd80f",
//     columnIndex: 3,
//     rowIndex: 18,  // Add this line with the starting row number
//     sheetName: "Sheet1" // Add sheet name
//   };

//   pollStatus();
// }


function handleLogin() {
  const subscriptionKey = document.getElementById("subscription-key").value;
  const apiKey = document.getElementById("api-key").value;

  if (!subscriptionKey || !apiKey) {
    AddtoLog("Error: Both Subscription Key and API Key are required.");
    return;
  }

localStorage.setItem("subscriptionKey", subscriptionKey);
localStorage.setItem("apiKey", apiKey);

document.getElementById("subscription-key").value = "";
document.getElementById("api-key").value = "";

const runButton = document.getElementById("run");
runButton.disabled = false;
runButton.classList.remove("disabled-button");

AddtoLog("Keys saved successfully.");
  fetchWorkerList(subscriptionKey, apiKey);
}

function handleLogout() {
  // Commented out to retain API key details between sessions
  localStorage.removeItem("subscriptionKey");
  localStorage.removeItem("apiKey");

  AddtoLog("Logged out successfully.");
  document.getElementById("worker-section").style.display = "none";

  // Gray out the submit button
  const runButton = document.getElementById("run");
  runButton.disabled = true;
  runButton.classList.add("disabled-button");

  // Make login details reappear
const loginFrame = document.getElementById("login-frame");
loginFrame.style.display = "block";

document.getElementById("subscription-key").value = "";
document.getElementById("api-key").value = "";
}

function fetchWorkerList(subscriptionKey, apiKey) {
    const myHeaders = new Headers();
    myHeaders.append("Ocp-Apim-Subscription-Key", subscriptionKey);
    myHeaders.append("API_KEY", apiKey);

    const requestOptions = {
        method: "GET",
        headers: myHeaders,
        redirect: "follow"
    };

    fetch("https://fn-doozer-py-05.azurewebsites.net/api/worker?view_type=lite", requestOptions)
        .then((response) => {
            if (!response.ok) {
                throw new Error(`Failed to fetch worker list. Status: ${response.status}`);
            }
            return response.json();
        })
        .then((result) => {
            const dropdown = document.getElementById("worker-dropdown");
            dropdown.innerHTML = ""; // Clear existing options
            
            AddtoLog("Debug: Worker list response - " + JSON.stringify(result.workers));
            
            if (result && Array.isArray(result.workers)) {
                result.workers.forEach((worker) => {
                    const option = document.createElement("option");
                    option.value = worker.WorkerID;
                    option.textContent = `${worker.Name}`;
                    option.setAttribute("data-picture", worker.Picture);
                    dropdown.appendChild(option);
                });
                
                dropdown.addEventListener("change", (event) => {
                    const selectedWorker = result.workers.find(worker => worker.WorkerID == event.target.value);
                    if (selectedWorker) {
                        displayAgentInfo(selectedWorker);
                    }
                });
            } else {
                AddtoLog("Error: Worker list is not properly structured. Response: " + JSON.stringify(result));
            }
            
            document.getElementById("worker-section").style.display = "flex";
            document.getElementById("run").disabled = false;
            
            // Hide login section after successful login
            document.getElementById("login-frame").style.display = "none";
            
            AddtoLog("Worker list fetched successfully.");
        })
        .catch((error) => {
            AddtoLog(`Error fetching worker list: ${error.message}`);
        });
}

function displayAgentInfo(agent) {
    // Create agent-info section if it doesn't exist
    let infoSection = document.getElementById("agent-info");
    if (!infoSection) {
        infoSection = document.createElement("div");
        infoSection.id = "agent-info";
        infoSection.style.marginTop = "20px";
        infoSection.style.padding = "10px";
        infoSection.style.border = "1px solid #ccc";
        infoSection.style.borderRadius = "5px";
        document.getElementById("app-body").appendChild(infoSection);
    }
    
    infoSection.innerHTML = ""; // Clear existing content

    const img = document.createElement("img");
    img.src = agent.Picture;
    img.alt = agent.Name;
    img.className = "agent-picture-large";
    img.style.width = "150px";
    img.style.height = "150px";
    img.style.borderRadius = "50%";
    img.style.display = "block";
    img.style.margin = "0 auto 15px auto";

    const name = document.createElement("h2");
    name.textContent = agent.Name;
    name.style.textAlign = "center";
    name.style.margin = "10px 0";
    name.style.color = "#333";

    const role = document.createElement("p");
    role.textContent = `Role: ${agent.Role}`;
    role.style.fontWeight = "bold";
    role.style.margin = "5px 0";

    const description = document.createElement("p");
    description.textContent = `Description: ${agent.Description}`;
    description.style.margin = "5px 0";

    const traits = document.createElement("p");
    traits.textContent = `Traits: ${agent.Traits}`;
    traits.style.margin = "5px 0";

    const skills = document.createElement("p");
    skills.textContent = `Skills: ${agent.Skills}`;
    skills.style.margin = "5px 0";

    const email = document.createElement("p");
    email.textContent = `Email: ${agent.Email}`;
    email.style.margin = "5px 0";

    infoSection.appendChild(img);
    infoSection.appendChild(name);
    infoSection.appendChild(role);
    infoSection.appendChild(description);
    infoSection.appendChild(traits);
    infoSection.appendChild(skills);
    infoSection.appendChild(email);
}

function clearLog() {
  const logWindow = document.getElementById("log-window");
  if (logWindow) {
    logWindow.innerHTML = "";
  }
}

function clearJobs() {
  

    try {
      data_dictionary = {};
    AddtoLog(`Cleared all jobs`);
  } catch (error) {
    AddtoLog(`${error}`);
  }

}


function status() {
  AddtoLog(JSON.stringify(data_dictionary));

}

export async function SubmitToAgent() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      // Validate that a range is selected
      if (!range) {
        AddtoLog("Error: No range selected. Please select a range in Excel and try again.");
        return;
      }

      try {
        // Load the range address and values
        range.load(["address", "columnIndex", "rowIndex", "values"]);
        await context.sync();
      } catch (error) {
        AddtoLog(`Error: Failed to load range or sync context. Details: ${error.message}`);
        throw error;
      }

      // Extract sheet name from range address
      let sheetName = "";
      const addressMatch = range.address.match(/'?([^'!]+)'?!/);
      if (addressMatch && addressMatch[1]) {
        sheetName = addressMatch[1];
      } else {
        AddtoLog("Error: Could not extract sheet name from range address.");
        return;
      }

      // Get the selected worker
      const workerDropdown = document.getElementById("worker-dropdown");
      const selectedWorkerId = workerDropdown.value;
      if (!selectedWorkerId) {
        AddtoLog("Error: No worker selected. Please select a worker from the dropdown.");
        return;
      }
      
      // Get the run button element
      const runButton = document.getElementById("run");
      runButton.textContent = "Processing...";
      const spinner = document.createElement("span");
      spinner.className = "spinner";
      spinner.style.marginLeft = "10px";
      spinner.style.border = "2px solid #f3f3f3";
      spinner.style.borderTop = "2px solid #0078d7";
      spinner.style.borderRadius = "50%";
      spinner.style.width = "16px";
      spinner.style.height = "16px";
      spinner.style.animation = "spin 1s linear infinite";
      runButton.appendChild(spinner);
      runButton.disabled = true;

      try {
        const progressBar = document.getElementById("progress-bar");
        if (!progressBar) {
          AddtoLog("Error: Progress bar element not found.");
          return;
        }

        progressBar.value = 0;
        progressBar.className = "";
        progressBar.classList.add("progress-waiting");

        const jsonData_string = JSON.stringify(range);

        const prompt = document.getElementById("prompt").value;

        AddtoLog(prompt);
        // Get the selected agent name from the dropdown list
        const workerDropdown = document.getElementById("worker-dropdown");
        const selectedAgentName = workerDropdown.options[workerDropdown.selectedIndex].textContent;

        // Prepare the payload for the WebService
        const payload = {
          workflow_short_name: "Agent Submit Excel",
          callback_url: "",
          doozer_name: selectedAgentName,
          variables: [
            {
              excelName: "None",
              jsonData: jsonData_string,
              guidance_prompt: prompt,
            },
          ],
        };

        const payloadString = JSON.stringify(payload);

        // Send the data to the WebService
        const response = await fetch("https://fn-doozer-py-05.azurewebsites.net/api/Queue", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Ocp-Apim-Subscription-Key": localStorage.getItem("subscriptionKey"),
            "API_KEY": localStorage.getItem("apiKey")
          },
          body: JSON.stringify(payload)
        });

        if (!response.ok) {
          throw new Error(`Failed to queue task. Status: ${response.status}`);
        }

        const result = await response.json();
        const instance_id = result.instance_id;

        // Store the instance ID and range info in the data dictionary
        data_dictionary[range.address] = {
          status: "pending",
          instance_id: instance_id,
          columnIndex: range.columnIndex,
          rowIndex: range.rowIndex,
          sheetName: sheetName
        };

        AddtoLog(`Task queued successfully. Instance ID: ${instance_id}`);
      } catch (error) {
        AddtoLog(`Error: Failed to queue task. Details: ${error.message}`);
      }
    });
  } catch (error) {
    AddtoLog(`Error: ${error.message}`);
  }
}
