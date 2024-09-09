var organizationURI = "https://org5d764994.crm11.dynamics.com";
var tenant = "qsolutions670.onmicrosoft.com";
var clientId = "aeb1c4ae-6990-4b20-8975-2fe6af68d4cc";
var pageUrl = "http://localhost:5500/Task5.html";
var user, authContext;

var apiVersion = "/api/data/v9.2/";
var endpoints = {
  orgUri: organizationURI,
};

window.config = {
  tenant: tenant,
  clientId: clientId,
  postLogoutRedirectUri: pageUrl,
  endpoints: endpoints,
  cacheLocation: "localStorage",
};
document.onreadystatechange = function () {
  if (document.readyState == "complete") {
    authenticate();
    if (!user) {
      authContext.login();
    }
  }
};

function authenticate() {
  authContext = new AuthenticationContext(config);

  var isCallback = authContext.isCallback(window.location.hash);
  if (isCallback) {
    authContext.handleWindowCallback();
  }

  var loginError = authContext.getLoginError();
  if (isCallback && !loginError) {
    window.location = authContext._getItem(
      authContext.CONSTANTS.STORAGE.LOGIN_REQUEST
    );
  } else {
    console.log(loginError);
  }

  user = authContext.getCachedUser();
}

var isCancelled = false;
var statusText = document.querySelector(".dialog__status");
var progressBar = document.querySelector(".dialog__progress-bar-value");

function normalizeDate(date) {
  return new Date(date.setHours(0, 0, 0, 0));
}

function startProcessing() {
  isCancelled = false;

  var startDate = normalizeDate(
    new Date(document.getElementById("startDate").value)
  );
  var endDate = normalizeDate(
    new Date(document.getElementById("endDate").value)
  );
  var today = normalizeDate(new Date());
  var maxEndDate = normalizeDate(new Date());
  maxEndDate.setDate(maxEndDate.getDate() + 7);

  if (
    startDate.getTime() !== today.getTime() ||
    startDate > endDate ||
    endDate > maxEndDate
  ) {
    alert("Invalid date range. Please select a valid range.");
    resetProgressBar();
    return;
  }

  authContext.acquireToken(organizationURI, (error, token) => {
    if (error || !token) {
      console.error("ADAL error occurred: " + error);
      return;
    }

    retrieveContacts(token, startDate, endDate);
  });
}

function cancelProcessing() {
  if (confirm("Are you sure that you want to cancel the processing?")) {
    isCancelled = true;
  }
}

function retrieveContacts(token, startDate, endDate) {
  var contactsQuery =
    "contacts?$select=firstname,lastname,birthdate,emailaddress1&$filter=tasks_last_birthday_congrat eq null";

  var req = new XMLHttpRequest();

  req.open(
    "GET",
    encodeURI(organizationURI + apiVersion + contactsQuery),
    true
  );

  req.setRequestHeader("Authorization", "Bearer " + token);
  req.setRequestHeader("Accept", "application/json");
  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  req.setRequestHeader("OData-MaxVersion", "4.0");
  req.setRequestHeader("OData-Version", "4.0");

  req.onreadystatechange = function () {
    if (this.readyState == 4 /* complete */) {
      req.onreadystatechange = null;
      if (this.status == 200) {
        var result = JSON.parse(this.response).value;

        var contacts = result.filter((contact) => {
          if (contact.birthdate) {
            var birthdate = new Date(contact.birthdate);
            var newBirthdate = new Date();
            newBirthdate.setMonth(birthdate.getMonth(), birthdate.getDate());

            return (
              normalizeDate(newBirthdate) >= normalizeDate(startDate) &&
              normalizeDate(newBirthdate) <= normalizeDate(endDate)
            );
          }
          return false;
        });

        if (contacts.length) {
          processContacts(contacts);
        } else {
          alert("There are no contacts with birthdays in this date range.");
        }
      } else {
        var error = JSON.parse(this.response).error;
        console.log(error.message);
        resetProgressBar();
      }
    }
  };
  req.send();
}

function processContacts(contacts) {
  var totalContacts = contacts.length;
  var processedContacts = 0;
  var sentEmails = 0;
  var failedEmails = 0;

  resetProgressBar(totalContacts);

  function processNextContact(index) {
    if (index >= contacts.length || isCancelled) {
      statusText.innerText = isCancelled
        ? `Processing is cancelled successfully. ${processedContacts} contacts are processed, ${sentEmails} congratulation e-mails were successfully sent (sending of ${failedEmails} e-mails failed).`
        : `${processedContacts} contacts are processed, ${sentEmails} congratulation e-mails were successfully sent (sending of ${failedEmails} e-mails failed).`;

      return;
    }

    var contact = contacts[index];

    var address = getAddressByGender(contact.gendercode);

    sendBirthdayGreeting(
      address,
      contact.emailaddress1,
      contact.firstname,
      contact.lastname
    )
      .then(() => {
        sentEmails++;
        updateContactLastCongratDate(contact.contactid);
      })
      .catch(() => {
        failedEmails++;
      })
      .finally(() => {
        processedContacts++;
        updateProgressBar(processedContacts, totalContacts);

        processNextContact(index + 1);
      });
  }

  processNextContact(0);
}

function sendBirthdayGreeting(address, email, firstName, lastName) {
  return new Promise((resolve, reject) => {
    if (isCancelled) {
      reject("Cancelled");
      return;
    }
    Email.send({
      Host: "smtp.elasticemail.com",
      Username: "gribovskijvladimir0@gmail.com",
      Password: "D4F7435DC5FA5535970D89E3F85D6AAA1323",
      To: email,
      From: "gribovskijvladimir0@gmail.com",
      Subject: "Happy Birthday",
      Body: `${address}${firstName} ${lastName}, alles Gute zum Geburtstag!`,
    })
      .then((message) => {
        console.log("Success: " + message);
        message === "OK" ? resolve(true) : reject("Failed to send email");
      })
      .catch((error) => {
        console.error("Error: " + error);
        reject(error);
      });
  });
}

function getAddressByGender(genderCode) {
  if (genderCode === 1) {
    return "Sehr geehrter Herr ";
  } else if (genderCode === 2) {
    return "Sehr geehrte Frau ";
  } else {
    return "";
  }
}

function updateContactLastCongratDate(contactId) {
  var req = new XMLHttpRequest();

  req.open(
    "PATCH",
    encodeURI(organizationURI + apiVersion + "contacts(" + contactId + ")"),
    true
  );

  req.setRequestHeader(
    "Authorization",
    "Bearer " + authContext.getCachedToken(organizationURI)
  );
  req.setRequestHeader("Accept", "application/json");
  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  req.setRequestHeader("OData-MaxVersion", "4.0");
  req.setRequestHeader("OData-Version", "4.0");

  var data = JSON.stringify({
    tasks_last_birthday_congrat: new Date(),
  });
  req.send(data);

  req.onreadystatechange = function () {
    if (this.readyState == 4) {
      req.onreadystatechange = null;
      if (this.status === 204) {
      } else {
        var error = JSON.parse(this.response).error;
        console.log("Error updating contact: " + error.message);
      }
    }
  };
}

function updateProgressBar(processedContacts, totalContacts) {
  var progress = (processedContacts / totalContacts) * 100;
  progressBar.style.width = `${progress}%`;
  statusText.innerText = `Processed contacts ${processedContacts} from ${totalContacts}.`;
}

function resetProgressBar(totalContacts = 0) {
  statusText.innerText = `Processed contacts 0 from ${totalContacts}.`;
  progressBar.style.width = "0";
}
