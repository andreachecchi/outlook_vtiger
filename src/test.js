const url = "https://crm.openia.it/webservice.php";
const auth = btoa("crm-openia:4NW@G3H'T/k5gavYNn-ebH87]");

const data = new URLSearchParams({
  operation: "login",
  username: "ws",
  accessKey: "323e2307ff0848841f13aeed7b0b2043",
  a: 1,
});

fetch(url, {
  method: "POST",
  headers: {
    Authorization: `Basic ${auth}`,
    "Content-Type": "application/x-www-form-urlencoded",
  },
  body: data,
})
  .then((response) => {
    console.log(`Status Code: ${response.status}`);
    return response.text();
  })
  .then((text) => console.log(text))
  .catch((error) => console.error("Error:", error));
