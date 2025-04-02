fetch("feats.json")
  .then(response => response.json())
  .then(data => console.log(data[0])) // Use the data
  .catch(error => console.error("Error loading database:", error));

