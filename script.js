const urIGAS = "https://script.google.com/macros/s/AKfycbxb54iTocoLK6RwX3T1RKvA22PWh1xGZkNReMtLBsmp0StTlWFtJwhgu6WGt6_OAcv-/exec";

function kirimData(data) {
    fetch(urlGAS, {
        method: "POST",
        mode: "no-cors",
        body:JSON.stringify(data)
    })
    .then(() => alert("Berhasil Terkirim ke google sheet!"))
    .catch(err => console.error("Error:", err));
}

