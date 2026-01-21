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

      result.d.results.forEach(function(file) {
        const fileExtension = file.Name.split('.').pop();

        var fileIcon = '';

        switch(fileExtension) {
          case 'rtf':
            fileIcon = 'https://cdn-icons-png.flaticon.com/128/8361/8361296.png';
            break;
          case 'txt':
            fileIcon = 'https://cdn-icons-png.flaticon.com/128/136/136539.png';
            break;
          case 'zip':
            fileIcon = 'https://cdn-icons-png.flaticon.com/128/337/337960.png';
            break;
          case 'jpeg':
            fileIcon = 'https://cdn-icons-png.flaticon.com/128/1824/1824880.png';
            break;
          case 'png':
            fileIcon = 'https://cdn-icons-png.flaticon.com/128/1824/1824880.png';
            break;
          case 'docx':
            fileIcon = 'https://cdn-icons-png.flaticon.com/128/8361/8361174.png';
            break;
          case 'xlsx':
            fileIcon = 'https://cdn-icons-png.flaticon.com/128/8361/8361467.png';
            break;
          case 'pdf':
            fileIcon = 'https://cdn-icons-png.flaticon.com/128/337/337946.png'
            break;
          default:
            fileIcon = 'https://cdn-icons-png.flaticon.com/128/2965/2965335.png'
        }

        const dateObj = new Date(file.TimeLastModified);
        const createdDate = dateObj.toLocaleDateString('en-US', {
          weekday: 'short', // ddd
          month: 'short',   // mmm
          day: 'numeric'    // D
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

  /*
  const [quotation, attribution] = [...block.children].map((c) => c.firstElementChild);
  const blockquote = document.createElement('blockquote');
  // decorate quotation
  quotation.className = 'quote-quotation';
  blockquote.append(quotation);
  // decoration attribution
  if (attribution) {
    attribution.className = 'quote-attribution';
    blockquote.append(attribution);
    const ems = attribution.querySelectorAll('em');
    ems.forEach((em) => {
      const cite = document.createElement('cite');
      cite.innerHTML = em.innerHTML;
      em.replaceWith(cite);
    });
  }
  block.innerHTML = '';
  block.append(blockquote);
  */
}