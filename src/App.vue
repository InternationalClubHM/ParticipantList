<template>
  <img src="@/assets/logo.png" alt="Logo" class="logo" />
  <label for="file-upload" class="file-upload-label" :class="fileUploaded ? 'upload-success' : ''">
    {{ fileUploaded?.name || 'Excel-Datei auswählen' }}
  </label>
  <input type="file" id="file-upload" @change="onFileUpdate" accept=".xlsx" />
  <button class="convert-button" @click="convertButtonClicked()">
    Teilnehmerliste in PDF umwandeln!
  </button>
  <p class="result-text" :style="{ color: resultText == successText ? 'lime' : 'red' }">
    {{ resultText }}
  </p>
</template>

<script lang="ts">
import { convertToPdf } from '@/assets/js/convertToPdf.ts'

export default {
  data() {
    return {
      fileUploaded: null as File | null,
      successText: 'PDF erstellt und heruntergeladen!',
      resultText: '',
    }
  },
  methods: {
    // Ran when someone updates a file or the file-upload input changes in a similar way
    onFileUpdate(event: Event) {
      this.resultText = ''

      const target = event.target as HTMLInputElement
      if (target.files && target.files.length > 0) {
        const file = target.files[0]

        if (!file.name.endsWith('.xlsx')) {
          this.resultText = 'Bitte wähle eine .xlsx (Excel) Datei aus.'
          return
        }

        this.fileUploaded = file
      } else {
        this.fileUploaded = null
      }
    },
    async convertButtonClicked() {
      if (!this.fileUploaded) {
        this.resultText = 'Bitte wähle eine Datei aus.'
        return
      }

      // Convert the participant list to pdf
      // If, by any way, the file is not a .xlsx file, an error will be caught below.
      const resultPdf = await convertToPdf(this.fileUploaded).catch((e) => {
        // Print the error in any case
        console.log(e)

        // The error could be anything (In the best case just a string).
        if (typeof e === 'string') {
          this.resultText = e
        } else if (e instanceof Error) {
          this.resultText = e.message
        } else {
          this.resultText = 'Irgendwas hat nicht geklappt....'
        }
      })

      // There is no result if an error was thrown
      if (!resultPdf) return

      // If successful, immediately download the newly created pdf.
      const fileUrl = window.URL.createObjectURL(resultPdf)
      const alink = document.createElement('a')
      alink.href = fileUrl
      alink.download = 'Teilnehmerliste.pdf'
      alink.click()

      // Signalize success!
      this.resultText = this.successText
    },
  },
}
</script>
