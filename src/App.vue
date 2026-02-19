<template>
  <div class="container" @dragenter.prevent="dragActive = true">
    <Transition>
      <div
        v-if="dragActive"
        class="drag-window"
        @dragleave.prevent="dragActive = false"
        @drop.prevent="onDrop"
      >

        <p class="drag-text">Leg deine xlsx-Datei hier ab.</p>
      </div>
    </Transition>


    <img src="@/assets/logo.png" alt="Logo" class="logo" />
    <label
      :class="fileUploaded ? 'upload-success' : ''"
      class="file-upload-label"
      for="file-upload"
    >
      {{ fileUploaded?.name || 'Teilnehmerliste (Excel-Datei) auswählen' }}
    </label>
    <input id="file-upload" accept=".xlsx" type="file" @change="onFileUpdate"/>
    <button class="convert-button" @click="convertSubmit()">
      Liste in PDF umwandeln!
    </button>
    <p :style="{ color: resultText === successText ? 'lime' : 'red' }" class="result-text">
      {{ resultText }}
    </p>
  </div>
</template>

<script lang="ts" setup>
import {onMounted, onUnmounted, ref} from 'vue'
import { convertToPdf } from '@/assets/js/convertToPdf.ts'

// For the drag-and-drop action to work
function preventDefaults(e: Event) {
  e.preventDefault()
}

const events = ['dragenter', 'dragover', 'dragleave', 'drop']
const successText = "PDF erstellt und heruntergeladen!"

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

const fileUploaded = ref<File | null>(null)
const resultText = ref("");
const dragActive = ref(false);

// Ran when someone updates a file or the file-upload input changes in a similar way
const onFileUpdate = (event: Event) => {
  const target = event.target as HTMLInputElement
  if (target.files && target.files.length > 0) {
    const file = target.files[0]

    if (!file.name.endsWith('.xlsx')) {
      resultText.value = "Bitte wähle eine .xlsx (Excel) Datei aus."
      return
    }

    fileUploaded.value = file

    // Automatically simulate button click
    convertSubmit()
  } else {
    fileUploaded.value = null
  }
}

const convertSubmit = async () => {
  if (!fileUploaded.value) {
    resultText.value = 'Bitte wähle eine Datei aus.'
    return
  }

  // Convert the participant list to pdf
  // If, by any way, the file is not a .xlsx file, an error will be caught below.
  const resultPdf = await convertToPdf(fileUploaded.value).catch((e) => {
    // Print the error in any case
    console.log(e)

    // The error could be anything (In the best case just a string).
    if (typeof e === 'string') {
      resultText.value = e
    } else if (e instanceof Error) {
      resultText.value = e.message
    } else {
      resultText.value = 'Irgendwas hat nicht geklappt....'
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
  resultText.value = successText
}

const onDrop = (e: DragEvent) => {
  dragActive.value = false

  // Something didn't work quite well
  if (!e.dataTransfer?.files) {
    resultText.value = 'Die Drag-and-Drop Aktion hat leider nicht geklappt...'
    return
  }

  // Multiple files have been selected
  if (e.dataTransfer.files.length != 1) {
    resultText.value = 'Bitte ziehe nur eine Datei rein.'
    return
  }

  const fileDropped = e.dataTransfer.files[0]

  // The file has the wrong ending/type
  if (!fileDropped.name.endsWith('.xlsx')) {
    resultText.value = 'Bitte wähle eine .xlsx (Excel) Datei aus.'
    return
  }

  // Successfully uploaded
  fileUploaded.value = fileDropped

  // Automatically simulate button click
  convertSubmit()
}
</script>
