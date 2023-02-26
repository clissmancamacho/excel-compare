// polyfills required by exceljs
require("core-js/modules/es.promise")
require("core-js/modules/es.string.includes")
require("core-js/modules/es.object.assign")
require("core-js/modules/es.object.keys")
require("core-js/modules/es.symbol")
require("core-js/modules/es.symbol.async-iterator")
require("regenerator-runtime/runtime")
const ExcelJS = require("exceljs/dist/es5")
const express = require("express")
const app = express()
const port = 3000

const filename1 = "comparar.xlsx"

app.get("/", async (req, res) => {
  const workbook = new ExcelJS.Workbook()
  try {
    await workbook.xlsx.readFile(filename1)
    const worksheet = workbook.getWorksheet("PAGO DIARIO")
    const c1 = worksheet.getColumn(1)
    const c2 = worksheet.getColumn(2)
    const c3 = worksheet.getColumn(3)

    const codesBoleta = []
    let codesBoletaCompare = []
    const codesCompCredito = []
    let codesCompCreditoCompare = []
    const codesCV = []
    let codesCVCompare = []
    const codesFactura = []
    let codesFacturaCompare = []

    searchCodesAndPush(c1, codesBoleta, codesCompCredito, worksheet)
    searchCodesAndPush(c2, codesCV, codesCompCredito, worksheet)
    searchCodesAndPush(c3, codesFactura, codesCompCredito, worksheet)

    const worksheet2 = workbook.getWorksheet("CIERRE DE CAJA")
    worksheet2.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      const rows = row.values
      if (rowNumber !== 1) {
        if (rowNumber === 2) {
          codesBoletaCompare = rows
            .map((item, index) => {
              if (index !== 1) {
                const codeUnformatted = item.split(" ")
                let code = codeUnformatted[3]
                let amount = codeUnformatted[4]
                code = code?.replace("$", "")
                amount = amount?.replace(/\./g, "")
                return {
                  code: parseInt(code, 10),
                  amount: parseInt(amount, 10),
                }
              }
            })
            .filter((item) => item !== undefined)
        }
        if (rowNumber === 3) {
          codesCVCompare = rows
            .map((item, index) => {
              if (index !== 1) {
                const codeUnformatted = item.split(" ")
                let code = codeUnformatted[4]
                let amount = codeUnformatted[5]
                code = code?.replace("$", "")
                amount = amount?.replace(/\./g, "")
                return {
                  code: parseInt(code, 10),
                  amount: parseInt(amount, 10),
                }
              }
            })
            .filter((item) => item !== undefined)
        }
        if (rowNumber === 4) {
          codesFacturaCompare = rows
            .map((item, index) => {
              if (index !== 1) {
                const codeUnformatted = item.split(" ")
                let code = codeUnformatted[4]
                let amount = codeUnformatted[5]
                code = code?.replace("$", "")
                amount = amount?.replace(/\./g, "")
                return {
                  code: parseInt(code, 10),
                  amount: parseInt(amount, 10),
                }
              }
            })
            .filter((item) => item !== undefined)
        }
        if (rowNumber === 9) {
          codesCompCreditoCompare = rows
            .map((item, index) => {
              if (index !== 1) {
                const codeUnformatted = item.split(" ")
                let code = codeUnformatted[6]
                let amount = codeUnformatted[7]
                code = code?.replace("$", "")
                amount = amount?.replace(/\./g, "")
                return {
                  code: parseInt(code, 10),
                  amount: parseInt(amount, 10),
                }
              }
            })
            .filter((item) => item !== undefined)
        }
      }
    })
    const missingBoleta = []
    const missingCV = []
    const missingFactura = []
    const missingCompCredito = []
    verifyIfExistAndSameAmount(codesBoletaCompare, codesBoleta, missingBoleta)
    verifyIfExistAndSameAmount(codesCVCompare, codesCV, missingCV)
    verifyIfExistAndSameAmount(
      codesFacturaCompare,
      codesFactura,
      missingFactura
    )
    verifyIfExistAndSameAmount(
      codesCompCreditoCompare,
      codesCompCredito,
      missingCompCredito
    )

    return res.send({
      missingBoleta,
      missingCV,
      missingFactura,
      missingCompCredito,
    })
  } catch (error) {
    console.log(error)
  }
})

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})

const searchCodesAndPush = (col, arrayToPush, codesCompCredito, worksheet) => {
  col.eachCell((c, index) => {
    if (index !== 1 && index !== 2) {
      let value = c.value
      if (value && value !== "") {
        value =
          typeof value === "string"
            ? parseInt(value.replace("\n", ""), 10)
            : value
        const objToPush = {
          code: value,
          amount: worksheet.getCell(`G${index}`).value,
        }
        arrayToPush.push(objToPush)
        if (c?.style?.fill?.fgColor?.argb === "FFFF0000") {
          codesCompCredito.push(objToPush)
        }
      }
    }
  })
}

const verifyIfExistAndSameAmount = (arrayCompare, arrayCodes, arrayToPush) => {
  arrayCompare.forEach((value, index) => {
    const infoBoleta = arrayCodes.find((item) => item.code === value.code)
    if (!infoBoleta || infoBoleta.amount !== value.amount) {
      console.log(`Hay un problema con la boleta ${value.code}`)
      arrayToPush.push({
        cierreCaja: value,
        pagoDiario: infoBoleta ?? "No existe en pago diario",
      })
    }
  })
}
