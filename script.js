const urIGAS = "https://script.google.com/macros/s/AKfycbxrXPx9zXRgd91OL8Kke2ry6PmIvAc5aAoAauU8fmHO2m5gemC7lU1EDdi0bi55Mx6k/exec";

function kirimData(data) {
    fetch(urlGAS, {
        method: "POST",
        mode: "no-cors",
        body:JSON.stringify(data)
    })
    .then(() => alert("Berhasil Terkirim ke google sheet!"))
    .catch(err => console.error("Error:", err));
}
