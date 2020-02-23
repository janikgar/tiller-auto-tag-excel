/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Excel, Office, Map, Set */

class Rule {
  category: string;
  tag: Set<string>;
  description: string;
  account: string;
  institution: string;
  amount: string;

  constructor(category: string, tagRaw: string, description: string, account: string, institution: string, amount: string) {
    this.category = category;
    this.tag = new Set(tagRaw.split(","));
    this.tag.delete(""); // The request from Excel will parse in an empty cell as an empty string; we will remove this by default
    this.description = description;
    this.account = account;
    this.institution = institution;
    this.amount = amount;
  }

  matchDescription(testString: string) {
    return this.description === "" || testString.includes(this.description)
  }
  matchAccount(testString: string) {
    return this.account === "" || testString.includes(this.account)
  }
  matchInstitution(testString: string) {
    return this.institution === "" || testString.includes(this.institution)
  }
  matchAmount(testNumber: number) {
    let lowerBound = /(?:>([-\d.]+))/.exec(this.amount) || null; // The whole capture group, including "<" or ">", will match on the first returned item; below we will use lowerBound[1] to receive the capture group with just the number portion
    let upperBound = /(?:<([-\d.]+))/.exec(this.amount) || null;
    if (!lowerBound && !upperBound) {
      return true
    }
    if (lowerBound && upperBound) {
      return testNumber > parseFloat(lowerBound[1]) && testNumber < parseFloat(upperBound[1])
    }
    if (lowerBound) {
      return testNumber > parseFloat(lowerBound[1])
    }
    if (upperBound) {
      return testNumber < parseFloat(upperBound[1])
    }
  }
}

class Transaction {
  address?: number;
  date: Date;
  description: string;
  category: string;
  tags: Set<string>;
  amount: number;
  account: string;
  accountNum: string;
  institution: string;
  month: Date;
  week: Date;
  fullDescription: string;
  checkNumber: string;
  transactionId: string;
  categorized: boolean;
  constructor (rowArray: Array<string>, address: number = null) {
    this.date = new Date(rowArray[0]);
    this.description = rowArray[1];
    this.category = rowArray[2];
    this.tags = new Set(rowArray[3].split(","));
    this.tags.delete(""); // The request from Excel will return an empty cell as an empty string; removing this by default
    this.amount = parseFloat(rowArray[4].replace("$",""));
    this.account = rowArray[5];
    this.accountNum = rowArray[6];
    this.institution = rowArray[7];
    this.month = new Date(rowArray[8]);
    this.week = new Date(rowArray[9]);
    this.fullDescription = rowArray[10];
    this.checkNumber = rowArray[11];
    this.transactionId = rowArray[12];
    this.address = address;
    this.categorized = false; // This switch is false by default; it becomes true once categorized by a rule so the first rule takes precedence
  }
}

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = main;
  }
});

function logLocal(logString: string) {
  let par = document.getElementById("result-text");
  let textNode: HTMLParagraphElement = document.createElement("p");
  textNode.innerHTML = logString;
  par.appendChild(textNode);
}

function makeRules(values: string[][]): Array<Rule> {
  let outputData: Array<Rule> = [];
  values.shift(); // slide off the first row, which is column headers
  for (let row of values) {
    let outputRow = new Rule(row[0], row[1], row[2], row[3], row[4], row[5])
    outputData.push(outputRow)
  }
  return outputData
}

async function getAutoTagData() {
  let rules: Array<Rule>;
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem("AutoTag").getUsedRange(true);
    sheet.load();

    await context.sync()
      .then(() => {
        rules = makeRules(sheet.text);
      })
    })
  return rules
}

async function getTransactions(): Promise<Map<string, Transaction>> {
  let columnsMatched = ["Description", "Account", "Institution", "Amount", "Category", "Tags"];
  let txnShape: Object;
  await Excel.run(async context => {
    let sheet = context.workbook.worksheets.getItem("Transactions").getUsedRange(true);
    sheet.load("columnCount, rowCount");

    await context.sync()
      .then(() => {
        txnShape = {"columnCount": sheet.columnCount, "rowCount": sheet.rowCount}
      })
      .catch(err => {
        return err
      })
  })

  let txnPaginator: Map<string, Transaction> = new Map<string, Transaction>();
  let columnsToSearch: Object = {};
  await Excel.run(async context => {
    let sheet = context.workbook.worksheets.getItem("Transactions")
      .getRangeByIndexes(0, 0, 1, txnShape["columnCount"])
    sheet.load();
    await context.sync()
      .then(() => {
        for (let i = 0; i < sheet.values[0].length; i++) {
          let column = sheet.values[0][i];
          if (columnsMatched.includes(column)) {
            columnsToSearch[i] = column
          }
        }
      })
      .then(async () => {
        let step = 1000; // Number of transactions processed per batch
        for (let i = 1; i < txnShape["rowCount"]; i+= step) {
          let sheet = context.workbook.worksheets.getItem("Transactions")
            .getRangeByIndexes(i, 0, step, txnShape["columnCount"]).getUsedRangeOrNullObject(true);
          sheet.load();
    
          await context.sync()
            .then(() => {
              for (let j = 0; j < sheet.text.length; j++) {
                let row = sheet.text[j]
                let rowAddress = i + j + 1 // i = current step position; j = current row within this round-trip; +1 because of the header row
                let newRow = new Transaction(row, rowAddress);
                txnPaginator.set(newRow.transactionId, newRow);
              }
            })
            .catch(err => {
              logLocal(err.toString())
            })
        }
      })
    })
  return txnPaginator
}

async function runRules(rules: Array<Rule>, transactionObject: Map<string, Transaction>, uncategorizedOnly: boolean = false): Promise<Array<Transaction>> {
  let matches: Array<any> = [];
  let transactions = Array.from(transactionObject.values());
  logLocal(`Running with ${rules.length} rules and ${transactions.length} transactions`)
  logLocal(`Only running uncategorized transactions?: ${uncategorizedOnly}`)
  for (let transaction of transactions) {
    for (let rule of rules) {
      let ruleMatch = (
        rule.matchDescription(transaction.description) &&
        rule.matchAccount(transaction.account) &&
        rule.matchInstitution(transaction.institution) &&
        rule.matchAmount(transaction.amount)
      )
      let categorizable = (
        !transaction.categorized && // If already categorized earlier, don't re-categorize
        (
          !uncategorizedOnly || // If not already categorized and we're running all transactions 
          (
            uncategorizedOnly && // If not already categorized, we're only running uncategorized transactions, and the category is indeed empty
            transaction.category === ""
          )
        )
      )
      if (ruleMatch) {
        if (categorizable) {
          transaction.category = rule.category;
          transaction.categorized = true;
        }
        for (let tag of rule.tag) {
          transaction.tags.add(tag);
        }
      }
    }
    matches.push(transaction);
  }
  return matches
}

async function getTagColumns(): Promise<any> {
  let returnPromise = await Excel.run(async context => {
    let sheet = context.workbook.worksheets.getItem("Transactions");
    let searchRange = sheet.getRange();
    let tagSearch = searchRange.findOrNullObject("Tags", {
      completeMatch: true,
      matchCase: true,
    })
    tagSearch.load("columnindex");
    return context.sync().then(() => {
      if (tagSearch === null) {
        return
      }
      let tagColumn = String.fromCharCode(tagSearch.columnIndex + 65);
      return tagColumn
    })
  })
  return returnPromise
}

async function getTagLastRow(): Promise<any> {
  let returnPromise = await Excel.run(async context => {
    let sheet = context.workbook.worksheets.getItem("Transactions");
    let usedRange = sheet.getUsedRange(true);
    let usedRangeRow = usedRange.getLastRow();
    usedRangeRow.load("rowIndex");
    return context.sync()
      .then(() => {
        return usedRangeRow.rowIndex + 1;
      })
  })
  return returnPromise
}

function scrapeTags(txns: Array<Transaction>): string[][] {
  let tagList: Array<Array<string>> = [];
  for (let txn of txns) {
    let tagArray: string;
    if (txn.tags.size > 0) {
      tagArray = [...txn.tags].join(",")
    }
    tagList.push([tagArray]);
  }
  return tagList
}

async function rewriteTags(address: string, column: string[][]) {
  await Excel.run(async context => {
    let sheet = context.workbook.worksheets.getItem("Transactions");
    let rewriteRange = sheet.getRange(address);
    rewriteRange.values = column;
    return context.sync()
  })
}

async function main() {
  let autoTagData: Array<Rule> = [];
  let rewriteAddress: string;
  let tagColumn: number;
  let tagLastRow: number;
  await getTagColumns()
    .then(columnResult => {
      tagColumn = columnResult
      return getTagLastRow()
    })
    .then(rowResult => {
      tagLastRow = rowResult;
      return getAutoTagData()
    })
    .then(result => {
      autoTagData = result;
      return getTransactions()
    })
    .then(txnResult => {
      rewriteAddress = `Transactions!${tagColumn}2:${tagColumn}${tagLastRow}`
      return runRules(autoTagData, txnResult, false);
    })
    .then(ruleResult => {
      return scrapeTags(ruleResult);
    })
    .then(tagResult => {
      rewriteTags(rewriteAddress, tagResult);
    })
    .catch(err => { logLocal(err.toString());});
}