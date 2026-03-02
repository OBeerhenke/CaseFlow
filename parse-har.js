const fs = require('fs');

const harData = fs.readFileSync('har.json', 'utf8');
const har = JSON.parse(harData);

har.log.entries.forEach(entry => {
  if (entry.response.status === 400) {
    console.log("Found 400 response:");
    console.log("Body:", entry.response.content.text);
  }
});

