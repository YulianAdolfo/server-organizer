function disButton() {
    document.getElementById("btn_send").disabled = true
}
function analizeExtensionExcel() {
    let file = document.getElementById("file-excel").value  
    let getLengthFileName = file.length
    let getExtension = file.substring((getLengthFileName-1)-4)

    if(getExtension === ".xlsx") {
        ///server-file
        document.getElementById("btn_send").disabled = false
    }else {
        createAlert("Archivo no válido, compruebe que la extensión del archivo sea (.xslx)")
    }
}
function createAlert(message, color="red") {
    this.message = message
    this.color = color
    let divAlert = document.createElement("div")
    let informationContent = document.createElement("p")
    informationContent.innerHTML = this.message
    divAlert.appendChild(informationContent)
    divAlert.classList.add("alert-info")
    divAlert.style.backgroundColor = this.color

    setTimeout(() => {
       document.body.removeChild(document.body.lastChild) 
    }, 3000);

    document.body.appendChild(divAlert)
}
disButton()