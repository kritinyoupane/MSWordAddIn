const baseUrl = "http://127.0.0.1:8000";
var old_text = "";
const loginFormContainer = () => {
  // create login container element
  const loginContainer = document.createElement("div");
  loginContainer.id = "login";

  // create Navbar element
  const navbar = document.createElement("Navbar");
  loginContainer.appendChild(navbar);

  const logInContainer = document.createElement("div");
  logInContainer.className = "logInContainer";
  loginContainer.appendChild(logInContainer);

  // create login header element
  const loginHeader = document.createElement("div");
  loginHeader.className = "top";
  const loginTitle = document.createElement("h3");
  loginTitle.className = "ms-font-su";
  loginTitle.style.color = "teal";
  loginTitle.style.fontSize = "30px";
  loginTitle.style.fontWeight = "400";
  loginTitle.textContent = "Log In";
  loginHeader.appendChild(loginTitle);
  loginContainer.appendChild(loginHeader);

  // create login form element
  const loginForm = document.createElement("form");

  // create username input field
  const usernameInput = document.createElement("div");
  usernameInput.className = "formInput";
  const usernameLabel = document.createElement("label");
  usernameLabel.textContent = "Username";
  const usernameField = document.createElement("input");
  usernameField.type = "text";
  usernameField.placeholder = "john_doe";
  usernameField.id = "username";
  usernameInput.appendChild(usernameLabel);
  usernameInput.appendChild(usernameField);
  loginForm.appendChild(usernameInput);

  // create password input field
  const passwordInput = document.createElement("div");
  passwordInput.className = "formInput";
  const passwordLabel = document.createElement("label");
  passwordLabel.textContent = "Password";
  const passwordField = document.createElement("input");
  passwordField.type = "password";
  passwordField.placeholder = "";
  passwordField.id = "password";
  passwordInput.appendChild(passwordLabel);
  passwordInput.appendChild(passwordField);
  loginForm.appendChild(passwordInput);

  // create login and sign up buttons
  const buttonsContainer = document.createElement("div");
  buttonsContainer.className = "buttons";
  const loginButton = document.createElement("button");
  loginButton.id = "loginButton";
  loginButton.textContent = "Log In";
  loginButton.onclick = async (e) => {
    e.preventDefault();
    const myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");
    const raw = JSON.stringify({
      username: document.getElementById("username").value,
      password: document.getElementById("password").value,
    });
    const requestOptions = {
      method: "POST",
      body: raw,
      headers: myHeaders,
      redirect: "follow",
    };
    const response = await fetch(`${baseUrl}/auth/login`, requestOptions);
    const data = await response.json();
    console.log(data);
    if (data["loginStatus"]==="Login"){
      localStorage.setItem("token", data["token"]);
      localStorage.setItem("user", JSON.stringify(data["user"]));
      const mainContainer = document.getElementById("app-body");
      while (mainContainer.firstChild) {
        mainContainer.removeChild(mainContainer.firstChild);
      }
      const user = JSON.parse(localStorage.getItem("user"));
      if (user) {
        if (user.approvalStatus === "Approved") {
          mainContainer.appendChild(contentContainer());
        } else {
          mainContainer.appendChild(adminApprovalContainer());
        }
      }
    }
    else{

      displayErrorMessage({"error":"username or password is incorrect"});
    }
  };
  const signUpContainer = document.createElement("span");
  signUpContainer.id = "registration";
  const signUpTitle = document.createElement("h6");
  signUpTitle.textContent = "Haven't registered yet?";
  const signUpButton = document.createElement("button");
  signUpButton.id = "signupButton";
  signUpButton.textContent = "Sign Up";
  signUpButton.onclick = (e) => {
    e.preventDefault();
    const mainContainer = document.getElementById("app-body");
    while (mainContainer.firstChild) {
      mainContainer.removeChild(mainContainer.firstChild);
    }
    mainContainer.appendChild(signUpFormContainer());
  };
  signUpContainer.appendChild(signUpTitle);
  signUpContainer.appendChild(signUpButton);
  buttonsContainer.appendChild(loginButton);
  buttonsContainer.appendChild(document.createElement("br"));
  buttonsContainer.appendChild(document.createElement("br"));
  buttonsContainer.appendChild(signUpContainer);
  loginForm.appendChild(buttonsContainer);

  // add form to login container
  loginContainer.appendChild(loginForm);

  // add login container to document body
  return loginContainer;
};
const contentContainer = () => {
  // create parent div with id "content"
  const contentDiv = document.createElement("div");
  contentDiv.setAttribute("id", "content");

  // create div with id "cursorHighlight"
  const cursorHighlightDiv = document.createElement("div");
  cursorHighlightDiv.setAttribute("id", "cursorHighlight");
  cursorHighlightDiv.textContent = "Please select text to see question recommendations.";

  // create div with id "run" and class "ms-welcome__action ms-Button ms-Button--hero ms-font-xl"
  const runDiv = document.createElement("div");
  runDiv.setAttribute("id", "run");
  runDiv.setAttribute("class", "ms-welcome__action ms-Button ms-Button--hero ms-font-xl");
  runDiv.setAttribute("role", "button");
  runDiv.onclick = (e) => {
    e.preventDefault();
    run();
  };
  const runSpan = document.createElement("span");
  runSpan.setAttribute("class", "ms-Button-label");
  runSpan.textContent = "Start⤀";
  runDiv.appendChild(runSpan);

  // create div with id "display_question"
  const displayQuestionDiv = document.createElement("div");
  displayQuestionDiv.setAttribute("id", "display_question");
  const displayQuestionHeader = document.createElement("h4");
  displayQuestionHeader.textContent = "Similar Questions";
  const displayQuestionText = document.createElement("div");
  displayQuestionText.setAttribute("class", "text-primary");
  displayQuestionText.textContent = "Similar questions would be displayed here!";
  displayQuestionDiv.appendChild(displayQuestionHeader);
  displayQuestionDiv.appendChild(displayQuestionText);
  
  // append all child elements to parent div
  contentDiv.appendChild(cursorHighlightDiv);
  contentDiv.appendChild(runDiv);
  contentDiv.appendChild(displayQuestionDiv);
  contentDiv.appendChild(document.createElement("br"));

  const paraphraseContainer =document.createElement("div");
  paraphraseContainer.id="paraphraseQuestions";
  const displayParaphraseHeader = document.createElement("h4");
  displayParaphraseHeader.textContent = "Paraphrase Questions";
  paraphraseContainer.appendChild(displayParaphraseHeader);
  paraphraseContainer.style.cssText="display:none";
  contentDiv.appendChild(paraphraseContainer);

  const buttonsContainer = document.createElement("div");
  buttonsContainer.className = "buttons";
  const logoutButton = document.createElement("button");
  logoutButton.id = "logoutButton";
  logoutButton.textContent = "Log Out";
  logoutButton.onclick = (e) => {
    e.preventDefault();
    localStorage.removeItem("user");
    localStorage.removeItem("token");
    const mainContainer = document.getElementById("app-body");
    while (mainContainer.firstChild) {
      mainContainer.removeChild(mainContainer.firstChild);
    }
    mainContainer.appendChild(loginFormContainer());
  };
  buttonsContainer.appendChild(logoutButton);
  contentDiv.appendChild(buttonsContainer);

  // append parent div to the DOM
  return contentDiv;
};
const signUpFormContainer = () => {
  // create signup container element
  const signupContainer = document.createElement("div");
  signupContainer.id = "signup";

  // create Navbar element
  const navbar = document.createElement("Navbar");
  signupContainer.appendChild(navbar);

  // create registration container element
  const registrationContainer = document.createElement("div");
  registrationContainer.className = "registrationContainer";
  signupContainer.appendChild(registrationContainer);

  // create registration header element
  const registrationHeader = document.createElement("div");
  registrationHeader.className = "top";
  const registrationTitle = document.createElement("h3");
  registrationTitle.className = "ms-font-su";
  registrationTitle.style.color = "teal";
  registrationTitle.style.fontSize = "30px";
  registrationTitle.style.fontWeight = "400";
  registrationTitle.textContent = "Sign Up";
  const registrationSubtitle = document.createElement("h3");
  registrationSubtitle.style.fontSize = "15px";
  registrationSubtitle.textContent = "It's quick and easy";
  registrationHeader.appendChild(registrationTitle);
  registrationHeader.appendChild(registrationSubtitle);
  registrationContainer.appendChild(registrationHeader);

  // create horizontal rule element
  const hr = document.createElement("hr");
  registrationContainer.appendChild(hr);

  // create registration form element
  const registrationForm = document.createElement("form");
  // create username name input field
  const userNameInput = document.createElement("div");
  userNameInput.className = "formInput";
  const userNameLabel = document.createElement("label");
  userNameLabel.textContent = "Username";
  const userNameField = document.createElement("input");
  userNameField.type = "text";
  userNameField.placeholder = "John";
  userNameField.id = "username";
  userNameInput.appendChild(userNameLabel);
  userNameInput.appendChild(userNameField);
  registrationForm.appendChild(userNameInput);

  // create first name input field
  const firstNameInput = document.createElement("div");
  firstNameInput.className = "formInput";
  const firstNameLabel = document.createElement("label");
  firstNameLabel.textContent = "First Name";
  const firstNameField = document.createElement("input");
  firstNameField.type = "text";
  firstNameField.placeholder = "John";
  firstNameField.id = "firstname";
  firstNameInput.appendChild(firstNameLabel);
  firstNameInput.appendChild(firstNameField);
  registrationForm.appendChild(firstNameInput);

  // create last name input field
  const lastNameInput = document.createElement("div");
  lastNameInput.className = "formInput";
  const lastNameLabel = document.createElement("label");
  lastNameLabel.textContent = "Last Name";
  const lastNameField = document.createElement("input");
  lastNameField.type = "text";
  lastNameField.placeholder = "Doe";
  lastNameField.id = "lastname";
  lastNameInput.appendChild(lastNameLabel);
  lastNameInput.appendChild(lastNameField);
  registrationForm.appendChild(lastNameInput);

  // create email input field
  const emailInput = document.createElement("div");
  emailInput.className = "formInput";
  const emailLabel = document.createElement("label");
  emailLabel.textContent = "Email";
  const emailField = document.createElement("input");
  emailField.type = "email";
  emailField.id = "email";
  emailField.placeholder = "johndoe@gmail.com";
  emailInput.appendChild(emailLabel);
  emailInput.appendChild(emailField);
  registrationForm.appendChild(emailInput);

  // create password input fields
  const passwordInput = document.createElement("div");
  passwordInput.className = "formInput";
  const passwordLabel = document.createElement("label");
  passwordLabel.textContent = "Password";
  const passwordField = document.createElement("input");
  passwordField.type = "password";
  passwordField.id = "password";
  passwordField.placeholder = "";
  passwordInput.appendChild(passwordLabel);
  passwordInput.appendChild(passwordField);
  registrationForm.appendChild(passwordInput);

  // create confirm password input fields
  const confirmPasswordInput = document.createElement("div");
  confirmPasswordInput.className = "formInput";
  const confirmPasswordLabel = document.createElement("label");
  confirmPasswordLabel.textContent = "Confirm Password";
  const confirmPasswordField = document.createElement("input");
  confirmPasswordField.type = "password";
  confirmPasswordField.id = "confirmpassword";
  confirmPasswordField.placeholder = "";
  confirmPasswordInput.appendChild(confirmPasswordLabel);
  confirmPasswordInput.appendChild(confirmPasswordField);
  registrationForm.appendChild(confirmPasswordInput);

  // create register button
  const registerButton = document.createElement("button");
  registerButton.id = "registerButton";
  registerButton.textContent = "Register";
  registerButton.onclick = async (e) => {
    e.preventDefault();
    const myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");
    const raw = JSON.stringify({
      first_name: document.getElementById("firstname").value,
      last_name: document.getElementById("lastname").value,
      email: document.getElementById("email").value,
      username: document.getElementById("username").value,
      password: document.getElementById("password").value,
      password2: document.getElementById("confirmpassword").value,
    });
    const requestOptions = {
      method: "POST",
      body: raw,
      headers: myHeaders,
      redirect: "follow",
    };
    const response = await fetch(`${baseUrl}/auth/register`, requestOptions);
    const data = await response.json();
    if (!data["error"]){
      const mainContainer = document.getElementById("app-body");
      while (mainContainer.firstChild) {
        mainContainer.removeChild(mainContainer.firstChild);
      }
      mainContainer.appendChild(loginFormContainer());
    }
    else{
      displayErrorMessage(data["error"]);
    }
  };
  registrationForm.appendChild(registerButton);
  const backButton = document.createElement("div");
  backButton.className = "back";
  backButton.innerHTML = "&#8592";
  backButton.onclick = ()=>{
    const mainContainer = document.getElementById("app-body");
      while (mainContainer.firstChild) {
        mainContainer.removeChild(mainContainer.firstChild);
      }
      mainContainer.appendChild(loginFormContainer());
  }
  registrationForm.appendChild(backButton);
  // add form to registration container
  registrationContainer.appendChild(registrationForm);

  // add registration container to document body
  return signupContainer;
};
const displayErrorMessage=(messages)=>{
  while(document.getElementById("error-display").firstChild)
  {
    document.getElementById("error-display").removeChild(document.getElementById("error-display").firstChild)
  }
  const alertContainer = document.createElement("div");
  alertContainer.className = "alert";
  const closeContainer = document.createElement ("span")
  closeContainer.innerHTML="&times";
  closeContainer.className="closebtn";
  closeContainer.onclick = ()=>{
    while(document.getElementById("error-display").firstChild)
    {
      document.getElementById("error-display").removeChild(document.getElementById("error-display").firstChild)
    }
  };
  const olElement = document.createElement("ul");
  Object.keys(messages).forEach(key=>{
    const liElement = document.createElement("li");
    liElement.innerHTML=messages[key];
    olElement.appendChild(liElement);
  })
  alertContainer.appendChild(closeContainer);
  alertContainer.appendChild(olElement)
  document.getElementById("error-display").appendChild(alertContainer)
}
const adminApprovalContainer = () => {
  const adminApproval = document.createElement("div");
  adminApproval.id = "pending";
  const h5 = document.createElement("h5");
  h5.class = "ms-font-xl";
  h5.textContent = "Sorry! Your account still need admin approval.";
  adminApproval.appendChild(h5);
  adminApproval.appendChild(document.createElement("br"));
  const adminContent = document.createElement("p");
  adminContent.textContent = "To view this page please";
  adminContent.style.cssText = `display:inline`;
  adminApproval.appendChild(adminContent);
  const contactContent = document.createElement("p");
  contactContent.textContent = "Contact your Admin";
  contactContent.style.cssText = `font-weight: 600; color: teal;`;
  adminApproval.appendChild(contactContent);
  const logoutButton = document.createElement("button");
  logoutButton.id = "logoutButton";
  logoutButton.textContent = "Log Out";
  logoutButton.onclick = (e) => {
    e.preventDefault();
    localStorage.removeItem("user");
    localStorage.removeItem("token");
    const mainContainer = document.getElementById("app-body");
    while (mainContainer.firstChild) {
      mainContainer.removeChild(mainContainer.firstChild);
    }
    mainContainer.appendChild(loginFormContainer());
  };
  adminApproval.appendChild(logoutButton);
  return adminApproval;
};
document.addEventListener("DOMContentLoaded", () => {
  const user = JSON.parse(localStorage.getItem("user"));
  const token = localStorage.getItem("token");
  if (user && token) {
    const mainContainer = document.getElementById("app-body");
    while (mainContainer.firstChild) {
      mainContainer.removeChild(mainContainer.firstChild);
    }
    if (user.approvalStatus === "Approved") {
      mainContainer.appendChild(contentContainer());
    } else {
      mainContainer.appendChild(adminApprovalContainer());
    }
  } else {
    const mainContainer = document.getElementById("app-body");
    while (mainContainer.firstChild) {
      mainContainer.removeChild(mainContainer.firstChild);
    }
    mainContainer.appendChild(loginFormContainer());
  }
});
async function run() {
  return Word.run(async (context) => {
    setInterval(async function () {
      var paragraphx = context.document.getSelection();
      context.load(paragraphx);
      await context.sync();
      if (paragraphx.text !== "" && paragraphx.text !== old_text) {
        console.log(paragraphx.text);
        old_text = paragraphx.text;
        document.getElementById("cursorHighlight").innerHTML = paragraphx.text;
        const myHeaders = new Headers();
        myHeaders.append("Content-Type", "application/json");
        const tokenKey = localStorage.getItem("token");
        myHeaders.append("Authorization", `Token ${tokenKey}`);
        const raw = JSON.stringify({
          queryQuestion: paragraphx.text,
        });
        const requestOptions = {
          method: "POST",
          body: raw,
          headers: myHeaders,
          redirect: "follow",
        };
        const response = await fetch(`${baseUrl}/api/queryQuestion/`, requestOptions);
        const data = await response.json();
        const displayQuestion = document.getElementById("display_question");
        while (displayQuestion.firstChild) {
          displayQuestion.removeChild(displayQuestion.firstChild);
        }
        data.map((d) => {
          $("#display_question").append(
            '<div class="row">' +
              '<div class="col-sm-10 alert alert-primary questionset"  role="alert">' +
              '<div class="frame1">'+
              // ' '+'<button onclick=clone(this)>'+
              d.question +
              '</div>' +
              '<div class="frame2">' +
              '<i class="fas fa-clone" onclick="clone(this, \'' +
              d.question +
              "')\"></i>&nbsp;&nbsp;" + 
              '</div>' +
              // '<div class="hide_overflow">' +
              // response[key].name +
              // "</div>" +
              // "</div>" +
              // '<div class="alert alert-secondary" role="alert">' +
              // response[key].date +
              "</div>" +
              "</div>"
          );
        });
        const paraphraseButton = document.createElement("button");
        paraphraseButton.id = "paraphrase";
        paraphraseButton.textContent = "Paraphrase";
        paraphraseButton.onclick = async (e) => {
          e.preventDefault();
          console.log("Clicked");
          const data = await paraPhrase(old_text,3)
          const container = document.getElementById("paraphraseQuestions");
          container.style.cssText="display:block";
          console.log(data)
          while(container.firstChild)
            {
              container.removeChild(container.firstChild);
            }
          data["choices"].forEach(d => {
            const da=document.createElement("p");
            da.textContent=d.text;
            
            container.appendChild(da)
          });
        };
        displayQuestion.appendChild(paraphraseButton);
      }
    }, 1000);

    const clone= (event)=> {
      Word.run(async (context) => {
        // Get the selected paragraph
        const paragraphx = context.document.getSelection();
        context.load(paragraphx);
        await context.sync();
        // Get the text content of the parent element of the clicked button
        const value = event.target.parentElement.textContent;
        // Insert the text into the selected paragraph, replacing its contents
        paragraphx.insertText(value, Word.InsertLocation.replace);

        document.getElementById("run").style.display = "none";
        await context.sync();
      });
    }
    const paraPhrase = async (query, noOfResult = 1) => {
      var myHeaders = new Headers();
      myHeaders.append("Content-Type", "application/json");
      myHeaders.append("Authorization", "Bearer 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");
      const prompt = `Paraphrase "${query}"`
      // const token_length = Math.min( parseInt( query.length /4 +10), 128)

      var raw = JSON.stringify({
          "model": "text-davinci-003",
          "prompt": prompt,
          "max_tokens": 128,
          "temperature": 0.99,
          "n": noOfResult
      });

      var requestOptions = {
          method: 'POST',
          headers: myHeaders,
          body: raw,
          redirect: 'follow'
      };

      const result = await fetch("https://api.openai.com/v1/completions", requestOptions);
      const json = await result.json()
      return json;
  }
    // carry out the function here

    // insert a paragraph at the end of the document.
    // const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    document.getElementById("run").style.display = "none";
    await context.sync();
  });
}
