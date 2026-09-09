/* =========================================================
   ZYALO LOGIN
========================================================= */

const VALID_USERNAME = "Lakshay";
const VALID_PASSWORD = "Lakshay@25";


/* =========================================================
   LOGIN
========================================================= */

function login() {

    const username =
        document
            .getElementById("username")
            .value
            .trim();


    const password =
        document
            .getElementById("password")
            .value;


    const message =
        document.getElementById(
            "login-message"
        );


    if (
        username === VALID_USERNAME &&
        password === VALID_PASSWORD
    ) {

        localStorage.setItem(
            "zyaloLoggedIn",
            "true"
        );


        localStorage.setItem(
            "zyaloUser",
            username
        );


        window.location.href =
            "collection.html";


    } else {

        message.textContent =
            "Invalid username or password.";

    }

}



/* =========================================================
   CHECK AUTHENTICATION
========================================================= */

function checkAuthentication() {

    const loggedIn =
        localStorage.getItem(
            "zyaloLoggedIn"
        );


    if (loggedIn !== "true") {

        window.location.href =
            "index.html";

        return;

    }

}



/* =========================================================
   LOGOUT
========================================================= */

function logout() {

    localStorage.removeItem(
        "zyaloLoggedIn"
    );


    localStorage.removeItem(
        "zyaloUser"
    );


    window.location.href =
        "index.html";

}