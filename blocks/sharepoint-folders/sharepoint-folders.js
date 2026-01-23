export function jsx(html, ...args) {
    return html.slice(1).reduce((str, elem, i) => str + args[i] + elem, html[0]);
}

export default async function decorate(block) {
  var renderHTML = `
    <div class="sharepoint-folders-none"></div>
  `;
  const sharepointFolders = block.querySelector('a');

  // exit link found
  if(sharepointFolders) {
    const sharepoint_endpoint = 'https://defaultfa7b1b5a7b34438794aed2c178dece.e1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7cf1dd3c93b243bda49882aaf73aa9da/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QCJox8uNJ8MQGi9Ma-djzuk33tJXJ6NeTflW7gPboSg';
    const sharepointFolderUrl = sharepointFolders.getAttribute('href');
    //sharepointFolders.remove();
    const data = {
        "folder": sharepointFolderUrl
    };

    try {
      const response = await fetch(sharepoint_endpoint, {
        method: 'POST', // Specify the method
        headers: {
          'Content-Type': 'application/json' // Declare the type of content
        },
        body: JSON.stringify(data) // Convert the JavaScript object to a JSON string
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const result = await response.json(); // Parse the JSON response

      if(result.d.results.length > 0) {
        renderHTML = '';
      }

      const fileIconMap = {
        rtf:  'https://cdn-icons-png.flaticon.com/128/9496/9496453.png',
        txt:  'https://cdn-icons-png.flaticon.com/128/9496/9496434.png',
        zip:  'https://cdn-icons-png.flaticon.com/128/9496/9496565.png',
        jpeg: 'https://cdn-icons-png.flaticon.com/128/9496/9496440.png',
        png:  'https://cdn-icons-png.flaticon.com/128/9496/9496440.png',
        docx: 'https://cdn-icons-png.flaticon.com/128/9496/9496487.png',
        xlsx: 'https://cdn-icons-png.flaticon.com/128/9496/9496502.png',
        pptx: 'https://cdn-icons-png.flaticon.com/128/9496/9496492.png',
        pdf:  'https://cdn-icons-png.flaticon.com/128/9496/9496432.png'
      };
      const defaultIcon = 'https://cdn-icons-png.flaticon.com/128/2965/2965335.png';

      result.d.results.forEach(function(file) {
        const fileExtension = file.Name.split('.').pop();
        const fileIcon = fileIconMap[fileExtension] || defaultIcon;

        const dateObj = new Date(file.TimeLastModified);
        const createdDate = dateObj.toLocaleDateString('en-US', {
          month: 'short',   // mmm
          day: 'numeric',    // D
          year: 'numeric'  // YYYY
        })

        renderHTML += `
          <div class="sharepoint-folders-item">
            <div class="file-img">
              <img src="${fileIcon}" alt="File Icon Thumbnail"/>
            </div>
            <div class="file-info">
              <div class="file-name">${file.Name}</div>
              <div class="file-date">${createdDate}</div>
            </div>
            <div class="file-actions">
              <button class="download" target="_blank" onclick="window.open('${file.LinkingUri}', '_blank')">Download</button>
            </div>
          </div>`;
      });
    } catch (error) {
      console.error('Error:', error);
    }
  }

  block.innerHTML = renderHTML;
}