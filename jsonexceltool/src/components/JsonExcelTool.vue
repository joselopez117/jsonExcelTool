<script setup>
import { reactive, ref, onMounted } from "vue";
import { read, utils, writeFileXLSX } from "xlsx";

const rows = ref([]);
const rawJsonString = ref("");
const counter = ref(0);
const hasError = ref(false);

function getFlatObject(object) {
  function iter(o, p) {
    if (o && typeof o === "object") {
      Object.keys(o).forEach(function (k) {
        iter(o[k], p.concat(k));
      });
      return;
    }
    path[p.join(".")] = o;
  }

  var path = {};
  iter(object, []);
  return path;
}

//filter json then export to xlsx format
function jsonToExcelFile() {
  try {
    //parse string into a json object
    const jsonObject = JSON.parse(rawJsonString.value);
    //parse the json object before throwing it into the json_to_sheet method
    const parsedJsonObject = getFlatObject(jsonObject);
    const arrayObject = Object.keys(parsedJsonObject).map((key) => [
      key,
      parsedJsonObject[key],
    ]);
    const ws = utils.json_to_sheet(arrayObject);
    const wb = utils.book_new();
    utils.book_append_sheet(wb, ws, "Data");
    writeFileXLSX(wb, `JsonToExcel${counter.value}.xlsx`);
    hasError.value = false;
  } catch (error) {
    console.log(error);
    hasError.value = true;
  }
  counter.value++;
}
</script>

<template>
  <div class="json-excel-tool">
    <h1 class="header">JsonExcelTool</h1>
    <p class="hello-world">
      Hello world! This tool was made to convert jsons to excel and back
    </p>
    <div class="counter">{{ counter }}</div>
    <div class="text-and-buttons">
      <div class="text-and-buttons__text">
        <div v-if="hasError" class="error-overlay">
          Json is not in correct format. Try Again
        </div>
        <textarea
          class="json-text-area"
          placeholder="Paste your json here"
          v-model="rawJsonString"
        ></textarea>
      </div>
      <div class="text-and-buttons__buttons">
        <button @click="jsonToExcelFile">Convert Json to Excel</button>
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

  .counter {
    font-style: normal;
  }

  .text-and-buttons {
    position: relative;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    width: max-content;
  }

  .text-and-buttons__text {
    width: max-content;
    .json-text-area {
      justify-content: center;
      min-width: 20%;
      min-height: 10rem;
      max-height: 30rem;
      height: 10rem;
      width: 39.6rem;
      resize: vertical;
      textarea {
        &::placeholder {
          font-style: italic;
        }
      }
    }
  }

  .error-overlay {
    position: absolute;
    background-color: salmon;
    height: 2rem;
    border-radius: 0.5rem;
    width: 40rem;
    text-align: center;
    vertical-align: middle;
  }

  .text-and-buttons__buttons {
    display: flex;
    flex-direction: row;
    width: auto;
  }
}
</style>
