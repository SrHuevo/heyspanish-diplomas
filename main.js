function sendEmail(body, attachments, wildcards) {
    const options = {
        Host: document.getElementById('host').value,
        Username : document.getElementById('user').value,
        Password : document.getElementById('pass').value,
        To : replaceWildcards(document.getElementById('to').value, wildcards),
        From : document.getElementById('from').value,
        Subject : replaceWildcards(document.getElementById('subject').value, wildcards),
        Body : body,
        Attachments : attachments,
    }
    return Email.send(options)
}

let attachmentNumber = 0
function addAttachment() {
    collapseAccordions(true)
    attachmentNumber++
    const auxNode = document.createElement('div')
    auxNode.innerHTML = `
        <div id="attachment-${attachmentNumber}" class="accordion-item">
          <h2 class="accordion-header" id="heading-attachment-${attachmentNumber}">
            <button class="accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-attachment-${attachmentNumber}" aria-expanded="true" aria-controls="collapseOne">
              Attachment ${attachmentNumber}
            </button>
          </h2>
          <div id="collapse-attachment-${attachmentNumber}" class="accordion-collapse collapse show" aria-labelledby="heading-attachment-${attachmentNumber}" data-bs-parent="#accordion-attachments">
            <div class="accordion-body attachment-config">
                <div class="row">
                    <label for="attachment-${attachmentNumber}-oritentation" class="col-sm-2 col-form-label">Orientación</label>
                    <div class="col-sm-10">
                        <select id="attachment-${attachmentNumber}-oritentation" class="form-select" onchange="saveParameter(event);">
                            <option value="portrait">portrait</option>
                            <option value="landscape">landscape</option>
                        </select>
                    </div>
                </div>
                <div class="row">
                    <label for="attachment-${attachmentNumber}-name" class="col-sm-2 col-form-label">Nombre</label>
                    <div class="col-sm-10">
                        <input type="text" class="form-control" id="attachment-${attachmentNumber}-name" onchange="saveParameter(event);">
                    </div>
                </div>
                <textarea id="attachment-${attachmentNumber}-html" onchange="saveParameter(event);render(event)"></textarea>
                <embed id="attachment-${attachmentNumber}-place" type="application/pdf" style="width: 100%;">
            </div>
          </div>
        </div>
    `
    const element = auxNode.children[0]
    document.getElementById('accordion-attachments').appendChild(element)
    element.scrollIntoView();
}

let students = []
async function readXLS(event) {
    collapseAccordions(true)

    const file = event.target.files[0]
    const workbook = XLSX.read(await file.arrayBuffer());
    students = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]])

    const progressBar = document.getElementById('progress')
    progressBar.parentElement.hidden = false
    progressBar.ariaValueMax = students.length
    progressBar.innerHTML = '&nbsp;&nbsp;&nbsp;' + '0/' + students.length + '&nbsp;&nbsp;&nbsp;'
}

Promise.all(Object.entries(window.localStorage).sort(([id1], [id2]) => id1.localeCompare(id2)).map(async ([id, value]) => {
    try {
        if(/attachment-\d+-html/.test(id)) {
            addAttachment()
        }
        const target = document.getElementById(id)
        target.value = value
        if(/attachment-\d+-oritentation/.test(id)) {
            await render({target})
        } else if("email-textarea" === id) {
            onEmailChange({target})
        }    
    } catch(e) {
        console.error(e)
    }
})).then(() => {
    window.scrollTo(0, 0)
    collapseAccordions(true)
})
function collapseAccordions(close) {
    Array.from(document.querySelectorAll(`.accordion-button${close ? ':not(.collapsed)' : ''}`)).forEach(accordion => accordion.click())
}

function saveParameter(event) {
    window.localStorage[event.target.id] = event.target.value
}

async function sendEmails(event) {
    collapseAccordions(true)

    window.onbeforeunload = () => {
        return "Uff caramba parece que vas a cerrar la pagina con el proceso a medias";
    }
    event.target.disabled = true
    const progressBar = document.getElementById('progress')
    progressBar.classList.add('progress-bar-animated')
    progressBar.ariaValueNow = 0
    progressBar.style.width = '0%'
    progressBar.innerHTML = '&nbsp;&nbsp;&nbsp;' + '0/' + students.length + '&nbsp;&nbsp;&nbsp;'    

    for(let i = 0; i < students.length; i++) {
        const student = students[i]
        const mail = replaceWildcards(document.querySelector('#email-textarea').value, student)

        const attachments = await Promise.all(Array.from(document.querySelectorAll('.accordion-body.attachment-config')).map(async (attachmentBody) => {
            const orientation = attachmentBody.querySelector('select').value
            const name = attachmentBody.querySelector('input').value
            const template = attachmentBody.querySelector('textarea').value
            const html = replaceWildcards(template, student)
            const pdf = await generate(html, orientation)

            return { data: pdf, name }
        }))

        await sendEmail(mail, attachments, student)

        const index = i + 1; 
        progressBar.ariaValueNow = index
        progressBar.style.width = (index * 100 / students.length) + '%'
        progressBar.innerHTML = '&nbsp;&nbsp;&nbsp;' + index + '/' + students.length + '&nbsp;&nbsp;&nbsp;'    
    }

    progressBar.classList.remove('progress-bar-animated')

    window.onbeforeunload = null
    event.target.disabled = false
    setTimeout(() => alert('Proceso finalizado :D'), 1000)
}

function replaceWildcards(text, wildcards) {
    const genericWildcards = [
        ['date', new Date().toLocaleDateString('es')],
    ]

    return Object.entries(wildcards).concat(genericWildcards).reduce((html, [wildcardName, wildcardValue]) => {
        return html.replace(new RegExp('\{\{' + wildcardName.trim() + '\}\}', 'g'), wildcardValue)
    }, text)
}


function convert(oldImag, callback) {
    return new Promise(resolve => {
        const img = new Image();
        img.onload = function(){
            resolve(img)
        }
        img.setAttribute('crossorigin', 'anonymous');
        img.src = oldImag.src;
    })
}
async function getBase64Image(img) {
    const newImg = await convert(img)
    const canvas = document.createElement("canvas");
    canvas.width = newImg.width;
    canvas.height = newImg.height;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(newImg, 0, 0);
    return canvas.toDataURL("image/png");
}
async function generate(html, orientation) {
    window.scrollTo(0,0)
    const auxElement = document.createElement('div')
    auxElement.innerHTML = html
    const element = auxElement.children[0]
    const imgs = Array.from(element.querySelectorAll('img'))
    await Promise.all(imgs.map(async (img)=> {
        const base64 = await getBase64Image(img) 
        img.src = base64
    }))

    const opt = {
        margin:       0.2,
        filename:     'myfile.pdf',
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2 },
        jsPDF:        { unit: 'in', format: 'a4', orientation }
    };
    const pdf = html2pdf().set(opt).from(element)
    return btoa(await pdf.outputPdf())
}

async function render(event) {
    const scroll = window.scrollY
    const attachmentNumber = getAttachmentNumber(event.target)
    const html = document.getElementById('attachment-' + attachmentNumber + '-html').value
    const orientation = document.getElementById('attachment-' + attachmentNumber + '-oritentation').value
    const pdfBase64 = await generate(html, orientation)
    const place = document.getElementById('attachment-' + attachmentNumber + '-place')
    place.src = 'data:application/pdf;base64,' + pdfBase64
    window.scrollTo(0,scroll)
}

function onEmailChange(event) {
    document.getElementById('email-place').innerHTML = event.target.value
}

function getAttachmentNumber(element) {
    return element.id.replace(/[^-]+-/, '').replace(/-.+/, '')
}