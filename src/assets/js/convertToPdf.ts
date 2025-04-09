// Imports here
import { read, utils } from 'xlsx'
import { PDFDocument, PDFFont, PDFPage } from 'pdf-lib'
import fontkit from '@pdf-lib/fontkit'

interface Participant {
  name: string
  phoneNumber: string
  country: string
  exchangeType: ExchangeType
}

enum ExchangeType {
  Erasmus,
  Other,
  Tutor,
}

/** What has to be found for every participant in the Excel file */
interface ValidParticipant {
  'First Name': string
  'Last Name': string
  'Phone Number': string
  'Country of Origin': string
  'Exchange Type': string
}

const basePdfDirectory = './base_pdf.pdf'

// The x-coordinates of the items in the pdf. Meaning where the texts should be placed.
const NAME_X = 67
const MOBILE_X = 230
const COUNTRY_X = 351
const ERASMUS_X = 467
const OTHER_EXCHANGE_X = 502.5
const TUTOR_X = 537.75

const AMOUNT_PARTICIPANTS_FIRST_PAGE = 8
const AMOUNT_PARTICIPANTS_OTHER_PAGES = 16

const TUTOR_LIST_X = 210
const TUTOR_LIST_Y = 468

// The y-coordinate of the start of the participant list on the first page
const PARTICIPANTS_FIRST_PAGE_Y = 284.87
// The y-coordinate of the start of the participant list on all other pages
const PARTICIPANTS_OTHER_PAGES_Y = 506.15

// The row height, meaning how much space should be between the participants in the pdf.
const ROW_HEIGHT = 27.41

const DEFAULT_FONT_SIZE = 11

/** Converts the participant list file to pdf. Returns the pdf file.
 * If something does not work or the file is incorrect, an error is thrown.
 */
export async function convertToPdf(excelFile: File): Promise<Blob> {
  const participantData = await getParticipants(excelFile)
  const pdfBlob = await createPdf(participantData)
  console.log('PDF Created!')

  return pdfBlob
}

/** Get the participants of an event from the Excel file. Returns an array of the participants.
 * @param excelFile the Excel file created by the Google-form with all information about the participants.
 * The file must include certain columns (see mandatoryProperties).
 * */
async function getParticipants(excelFile: File): Promise<Participant[]> {
  // Convert the file to an array buffer
  const arrayBuffer = await excelFile.arrayBuffer()

  // Read the file using Sheet.js and get the first sheet of the file
  const workbook = read(arrayBuffer)
  const sheet = workbook.Sheets[workbook.SheetNames[0]]

  // Get a list of json objects containing information about each row (each participant)
  const sheetRows: object[] = utils.sheet_to_json(sheet)

  // The properties which must be found for each participant
  const mandatoryProperties = [
    'First Name',
    'Last Name',
    'Phone Number',
    'Country of Origin',
    'Exchange Type'
  ]

  const allParticipants: Participant[] = []

  // Iterate through all rows
  sheetRows.forEach((participant, index) => {
    const modifiedParticipant = participant as { [key: string]: string }

    // Trim each value.
    for (const key in participant) {
      // Also convert every cell to string
      modifiedParticipant[key.trim()] = modifiedParticipant[key].toString()
    }

    // For every participant, every mandatory property must exist. If one does not, an error is thrown.
    mandatoryProperties.forEach((property) => {
      if (!modifiedParticipant.hasOwnProperty(property)) {
        throw new Error(
          `Datei ungültig: '${property}' existiert nicht für den ${index + 1}. Teilnehmer.`
        )
      }
    })

    // Cast is fine, as we just checked if every property exists (an error would have been thrown).
    const validParticipant = modifiedParticipant as unknown as ValidParticipant

    // The name is the first name + last name
    const name = validParticipant['First Name'] + ' ' + validParticipant['Last Name']

    const phoneNumber = validParticipant['Phone Number']
    const countryOfOrigin = validParticipant['Country of Origin']
    const exchangeTypeString = validParticipant['Exchange Type']

    // Convert the string exchange type to enum.
    let exchangeType: ExchangeType
    if (exchangeTypeString.startsWith('Erasmus')) {
      exchangeType = ExchangeType.Erasmus
    } else if (exchangeTypeString.startsWith('Other')) {
      exchangeType = ExchangeType.Other
    } else if (exchangeTypeString.startsWith('Tutor')) {
      exchangeType = ExchangeType.Tutor
    } else {
      // If the exchange type cannot be assigned, throw an error
      throw new Error(`Der Austauschtyp '${exchangeTypeString}' ist ungültig für ${name}!`)
    }

    // Create a new Participant object with the relevant data
    allParticipants.push({
      name: name,
      phoneNumber: phoneNumber,
      country: countryOfOrigin,
      exchangeType: exchangeType
    })
  })

  // Return the list of all participants
  return allParticipants
}

/** Creates a filled-out pdf of the participant list
 * */
async function createPdf(allParticipants: Participant[]) {
  // Read the pdf file and convert it to a PDFDocument of pdf-lib
  const existingPdfBytes = await fetch(basePdfDirectory).then((res) => res.arrayBuffer())
  const pdfDoc = await PDFDocument.load(existingPdfBytes)

  // Fetch the Ubuntu font
  const fontUrl = './Font-Nunito-Regular.ttf'
  const fontBytes = await fetch(fontUrl).then((res) => res.arrayBuffer())

  // Embed the Nunito Regular font
  pdfDoc.registerFontkit(fontkit)
  const font = await pdfDoc.embedFont(fontBytes)

  const pdfPages = pdfDoc.getPages()

  // Get a list of all Tutors
  const tutors = allParticipants.filter(
    (participant) => participant.exchangeType === ExchangeType.Tutor
  )
  // Get the participants on the first page, can also be less than AMOUNT_PARTICIPANTS_FIRST_PAGE if there aren't many participants for the event.
  const firstParticipants = allParticipants.slice(0, AMOUNT_PARTICIPANTS_FIRST_PAGE)
  // Write information on the first page of the pdf.
  writeFirstPdfPage(pdfPages[0], tutors, firstParticipants, font)

  // Participants not already written onto the first page
  const remainingParticipants = allParticipants.slice(AMOUNT_PARTICIPANTS_FIRST_PAGE)

  // Write the participants on all other pages
  const otherPagesWritten = writeOtherPdfPages(remainingParticipants, pdfPages.slice(1), font)

  // Remove every page which hasn't been written on. Backwards, because the number of pages changes after one deletion.
  for (let i = pdfPages.length - 1; i > otherPagesWritten; i--) {
    pdfDoc.removePage(i)
  }

  // Everything written. Now serialize the PDFDocument to bytes (a Uint8Array)
  const pdfBytes = await pdfDoc.save()
  return new Blob([pdfBytes])
}

/** The first page is special, meaning that the participant list is smaller and starts lower.
 * Also, the Tutor names are added to the Tutors title row.
 * */
function writeFirstPdfPage(
  firstPdfPage: PDFPage,
  tutors: Participant[],
  firstParticipants: Participant[],
  font: PDFFont
) {
  // Get the list of tutors to write down at the title bar
  const tutorNameList = tutors.map((tutor) => tutor.name).join(', ')

  // Write down the tutors at the title bar
  firstPdfPage.drawText(tutorNameList, {
    x: TUTOR_LIST_X,
    y: TUTOR_LIST_Y,
    font: font,
    size: DEFAULT_FONT_SIZE
  })

  // Add first participants
  addAllParticipantsToPdfPage(firstPdfPage, firstParticipants, PARTICIPANTS_FIRST_PAGE_Y, font)
}

/** Add all participants, which were not written onto the first pdf page, to the other pages.
 * @return how many pages have been written
 * */
function writeOtherPdfPages(
  allParticipantsNotOnFirstPage: Participant[],
  remainingPdfPages: PDFPage[],
  font: PDFFont
): number {
  let remainingParticipants = allParticipantsNotOnFirstPage
  let currentPage = 0

  // As long as there are remaining participants, add them
  while (remainingParticipants.length > 0) {
    // Throw an error if the pages are not enough
    if (currentPage > remainingPdfPages.length - 1)
      throw new Error(
        'Die originale PDF-Datei hat leider zu wenig Seiten. Bitte reduziere die Anzahl der Teilnehmer.'
      )

    // Write remaining participants (max AMOUNT_PARTICIPANTS_OTHER_PAGES)
    addAllParticipantsToPdfPage(
      remainingPdfPages[currentPage],
      remainingParticipants.slice(0, AMOUNT_PARTICIPANTS_OTHER_PAGES),
      PARTICIPANTS_OTHER_PAGES_Y,
      font
    )

    // Remove the just written participants from the remaining ones and add another page
    remainingParticipants = remainingParticipants.slice(AMOUNT_PARTICIPANTS_OTHER_PAGES)
    currentPage++
  }

  return currentPage
}

/** Writes the participants on a pdf page
 * @param pdfPage the page to write on
 * @param participantsToWrite all participants to write on the page. Don't have to fill the page
 * @param startY the y-coordinate of the first row
 * @param font the font to write the text in. Especially important for special characters.
 */
function addAllParticipantsToPdfPage(
  pdfPage: PDFPage,
  participantsToWrite: Participant[],
  startY: number,
  font: PDFFont
) {
  for (let i = 0; i < participantsToWrite.length; i++) {
    const yCoordinate = startY - ROW_HEIGHT * i
    addParticipantToPdf(pdfPage, participantsToWrite[i], yCoordinate, font)
  }
}

/** Writes down one participant of a pdf page.
 * @param pdfPage the page to write on
 * @param participant the participant which should be written
 * @param yCoordinate the y-coordinate of the row
 * @param font the font to write the text in. Especially important for special characters.
 * */
function addParticipantToPdf(
  pdfPage: PDFPage,
  participant: Participant,
  yCoordinate: number,
  font: PDFFont
) {
  function drawString(string: string, x: number) {
    // Write the participant's name
    pdfPage.drawText(string, {
      x: x,
      y: yCoordinate,
      font: font,
      size: DEFAULT_FONT_SIZE
    })
  }

  drawString(participant.name, NAME_X)
  drawString(participant.phoneNumber, MOBILE_X)
  drawString(participant.country, COUNTRY_X)

  // Figure out the position for the x as exchange type
  let exchangeTypeXPos = ERASMUS_X
  if (participant.exchangeType === ExchangeType.Other) exchangeTypeXPos = OTHER_EXCHANGE_X
  else if (participant.exchangeType === ExchangeType.Tutor) exchangeTypeXPos = TUTOR_X

  // Draw the x for the exchange type
  pdfPage.drawText('x', {
    x: exchangeTypeXPos,
    y: yCoordinate + 1,
    font: font,
    size: 17
  })
}
