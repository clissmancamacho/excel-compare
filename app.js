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

    searchCodesAndPush(c1, codesBoleta, codesCompCredito)
    searchCodesAndPush(c2, codesCV, codesCompCredito)
    searchCodesAndPush(c3, codesFactura, codesCompCredito)

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
                code = code?.replace("$", "")
                return parseInt(code, 10)
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
                code = code.replace("$", "")
                return parseInt(code, 10)
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
                code = code.replace("$", "")
                return parseInt(code, 10)
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
                code = code.replace("$", "")
                return parseInt(code, 10)
              }
            })
            .filter((item) => item !== undefined)
        }
      }
    })
    const missingBoleta = []
    codesBoletaCompare.forEach((value, index) => {
      if (!codesBoleta.find((item) => item === value)) {
        console.log(`El codigo de las boletas ${value} no está`)
        missingBoleta.push(value)
      }
    })
    const missingCV = []
    codesCVCompare.forEach((value, index) => {
      if (!codesCV.find((item) => item === value)) {
        console.log(`El codigo de las comprobantes de venta ${value} no está`)
        missingCV.push(value)
      }
    })
    const missingFactura = []
    codesFacturaCompare.forEach((value, index) => {
      if (!codesFactura.find((item) => item === value)) {
        console.log(`El codigo de las factura ${value} no está`)
        missingFactura.push(value)
      }
    })
    const missingCompCredito = []
    codesCompCreditoCompare.forEach((value, index) => {
      if (!codesCompCredito.find((item) => item === value)) {
        console.log(`El codigo de las factura ${value} no está`)
        missingCompCredito.push(value)
      }
    })
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

const searchCodesAndPush = (col, arrayToPush, codesCompCredito) => {
  col.eachCell((c, index) => {
    if (index !== 1 && index !== 2) {
      let value = c.value
      if (value && value !== "") {
        value =
          typeof value === "string"
            ? parseInt(value.replace("\n", ""), 10)
            : value
        arrayToPush.push(value)
        if (c?.style?.fill?.fgColor?.argb === "FFFF0000") {
          codesCompCredito.push(value)
        } 
      }
    }
  })
}
