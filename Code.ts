function getNewTransactions() {
  const sheetName = 'Transactions'
  const domainCol = 4
  const typeCol = 6;
  const checkBoxCol = 9
  const prevDataSheetName = 'TransLog'

  const ss = SpreadsheetApp.getActive()
  const tasksSheet = ss.getSheetByName(sheetName);
  const tasksRange = tasksSheet?.getRange('A2:J')
  let tasks = tasksRange?.getValues();
  tasks = tasks?.filter(row => row[typeCol] !== "")
  //tasks?.sort((a, b) => a[0] - b[0])

  const prevTasksRange = ss.getSheetByName(prevDataSheetName)?.getRange('A2:J')
  let prevTasks = prevTasksRange?.getValues()
  prevTasks = prevTasks?.filter(row => row[typeCol] !== "")
  //prevTasks?.sort((a, b) => a[0] - b[0])

  //Check which tasks are new
  let newTasks: any[][] = []
  if (tasks && prevTasks && tasks.length > prevTasks.length) {
    const nrOfNewTasks = tasks.length - prevTasks.length
    newTasks = tasks.slice(-nrOfNewTasks)
  }
  newTasks = newTasks.filter((task) => task[checkBoxCol] === false)

  if (newTasks.length > 0) {
    emailNewTasks(newTasks)
  }
  //overwrite TransLog sheet to make sure on the next run the same new tasks won't be emailed
  const taskValues = tasksRange?.getValues()
  if (taskValues && taskValues.length > 0 && taskValues[0].length > 0) {
    prevTasksRange?.setValues(taskValues);
  } else {
    console.log(`something funky happened with the task lengths`)
  }
}
9
function emailNewTasks(tasks: any[][]) {
  const emailsColumn = SpreadsheetApp.getActive().getSheetByName('Data')?.getRange('C2:C').getValues();
  if (!emailsColumn) {
    console.log('no emails are found in the emails column');
    return
  }

  let emails = ''
  for (const email of emailsColumn) {
    if (email[0] && email[0] != '') {
      emails += email[0] + ','
    }
  }

  const subject = 'New domain tasks for RGF registered'

  const docLink = 'https://docs.google.com/spreadsheets/d/1XvugI0JWm02rl3zEGxyKjHchha9EdgEElmo_zgGKpuU/edit?usp=sharing'
  let email = `New Domain tasks have been added to the 'RGF-ESM Domein Mutaties' sheet;\n\n`

  let i = 1;
  for (const task of tasks) {
    const type = task[6];
    const fullDomain = `${task[4]}.${task[5].replace('.', '')}`
    const entity = task[7]
    email += `${i}.  ${type} for ${fullDomain} for legal entity ${entity}\n`
    i++
  }

  email += `\nSee the RGF-ESM Domein Mutaties sheet for more detailed information: ${docLink}`;

  MailApp.sendEmail(emails, subject, email)
}