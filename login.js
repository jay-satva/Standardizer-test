document.getElementById("submitForm").addEventListener("submit", function(e){
    e.preventDefault()
    const email = document.getElementById('email').value
    const password = document.getElementById('password').value

    if(email === "jay@satvasolutions.com" && password === "123"){
        localStorage.setItem("isLoggedIn", "true")
        window.location.href = "index.html"
    }
    else{
        alert("Invalid credentials")
    }
}) 