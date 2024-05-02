<script setup>
import { ref } from "vue";
</script>

<template>
  <div class="json-excel-tool">
    <h1 class="header">JsonExcelTool</h1>
    <p class="hello-world">
      Hello world! This tool was made to convert jsons to excel
    </p>
    <div class="text-and-buttons">
      <div class="text">
        <textarea class="json-text-area" placeholder="Paste your json here"></textarea>
      </div>
      <div class="buttons">
        <button>Convert Json to Excel</button>
        <button>Convert Excel to Json</button>
      </div>
    </div>
  </div>
</template>

<style scoped lang="scss">
.json-excel-tool {
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  height: 100%;
  width: 100%;

  .header {
    display: flex;
  }

  .hello-world {
    display: flex;
  }

  .text-and-buttons {
    position: relative;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    width: 20rem;

    .text {
      width: max-content;
      .json-text-area {
        justify-content: center;
        min-width: 20%;
        min-height: 10rem;
        max-height: 30rem;
        height: 10rem;
        width: 40rem;
        resize: vertical;
        textarea {
          &::placeholder {
            font-style: italic;
          }
        }
      }
    }
    .buttons {
      display: flex;
      flex-direction: row;
      width: auto;
    }
  }
}
</style>
