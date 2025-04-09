<template>
  <div class="container" @dragenter.prevent="dragActive = true">
    <div class="drag-window" v-if="dragActive"  @dragleave.prevent="dragActive = false" @drop.prevent="onDrop">
      <p class="drag-text">
        Leg deine xlsx-Datei hier ab.
      </p>
    </div>

    <img src="@/assets/logo.png" alt="Logo" class="logo" />
    <label
      for="file-upload"
      class="file-upload-label"
      :class="fileUploaded ? 'upload-success' : ''"
    >
      {{ fileUploaded?.name || 'Excel-Datei auswählen' }}
    </label>
    <input type="file" id="file-upload" @change="onFileUpdate" accept=".xlsx" />
    <button class="convert-button" @click="convertButtonClicked()">
      Teilnehmerliste in PDF umwandeln!
    </button>
    <p class="result-text" :style="{ color: resultText == successText ? 'lime' : 'red' }">
      {{ resultText }}
    </p>
  </div>
</template>

<script lang="ts">
import { convertToPdf } from '@/assets/js/convertToPdf.ts'
import { onMounted, onUnmounted } from 'vue'

export default {
  setup() {
    function preventDefaults(e: Event) {
      e.preventDefault()
    }

    const events = ['dragenter', 'dragover', 'dragleave', 'drop']

    // Override the default behaviour of opening a file on drag and drop
    onMounted(() => {
      events.forEach((eventName) => {
        document.body.addEventListener(eventName, preventDefaults)
      })
    })

    onUnmounted(() => {
      events.forEach((eventName) => {
        document.body.removeEventListener(eventName, preventDefaults)
      })
    })
  },
  data() {
    return {
      fileUploaded: null as File | null,
      successText: 'PDF erstellt und heruntergeladen!',
      resultText: '',
      dragActive: false
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

    onDrop(e: DragEvent) {
      this.dragActive = false;

      // Something didn't work quite well
      if (!e.dataTransfer?.files) {
        this.resultText = 'Die Drag-and-Drop Aktion hat leider nicht geklappt...'
        return
      }

      // Multiple files have been selected
      if (e.dataTransfer.files.length != 1) {
        this.resultText = 'Bitte ziehe nur eine Datei rein.'
        return
      }

      const fileDropped = e.dataTransfer.files[0]

      // The file has the wrong ending/type
      if (!fileDropped.name.endsWith('.xlsx')) {
        this.resultText = 'Bitte wähle eine .xlsx (Excel) Datei aus.'
        return
      }

      // Successfully uploaded
      this.fileUploaded = fileDropped
    }
  }
}
</script>
