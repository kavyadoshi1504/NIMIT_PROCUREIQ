const loginForm = document.getElementById("loginForm");
const loginButton = document.getElementById("loginButton");
const message = document.getElementById("message");

loginForm.addEventListener("submit", async event => {
    event.preventDefault();

    const username = document.getElementById("username").value.trim();
    const password = document.getElementById("password").value;

    message.textContent = "";
    message.className = "message";
    loginButton.disabled = true;
    loginButton.textContent = "Signing in...";

    try {
        const response = await fetch("/login", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            credentials: "include",
            body: JSON.stringify({ username, password })
        });

        const data = await response.json();

        if (!response.ok) {
            message.textContent = data.message || "Login failed";
            message.className = "message error";
            return;
        }

        message.textContent = "Login successful. Redirecting...";
        message.className = "message success";

        window.location.href = data.redirect_url;
    } catch (error) {
        console.error("Login request failed:", error);
        message.textContent = "Unable to connect to the login service";
        message.className = "message error";
    } finally {
        loginButton.disabled = false;
        loginButton.textContent = "Login";
    }
});