/**
 * ============================================================
 * Fully Resumable File Tree Generator for "Archive and Backup"
 * ============================================================
 * Checkpoints after EVERY single file/folder processed, using
 * Drive's iterator continuation tokens - so no folder, no matter
 * how wide or deep, can cause a timeout to lose progress. Each
 * execution just does as much as it can in ~5 minutes, saves an
 * exact resume point, and reschedules itself.
 *
 * USAGE:
 *   1. Paste this whole file in, replacing any previous version.
 *   2. Run `startFileTreeGeneration` once (authorize Drive/Docs/
 *      Triggers when prompted).
 *   3. It self-schedules until the whole tree is written to a
 *      Google Doc. Just check back on the Doc - no manual
 *      re-running needed, even for huge/wide folders.
 *   4. `resetFileTreeGeneration()` wipes state to start over.
 * ============================================================
 */

const ROOT_FOLDER_ID = '0ANAtVEuavKC6Uk9PVA';
const PROCESS_FN_NAME = 'processNextFolder';
const TIME_BUDGET_MS = 5 * 60 * 1000; // 5 min - buffer under the 6 min cap
const TRIGGER_DELAY_MS = 60 * 1000;   // ~1 min between batches

/** One-time setup. Safe to call again - resumes if state already exists. */
function startFileTreeGeneration() {
  const props = PropertiesService.getScriptProperties();

  if (props.getProperty('DOC_ID')) {
    Logger.log('State already exists - resuming. Call resetFileTreeGeneration() first for a clean start.');
    processNextFolder();
    return;
  }

  const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const doc = DocumentApp.create(`${rootFolder.getName()} - File Tree`);
  const body = doc.getBody();
  body.appendParagraph(`📁 ${rootFolder.getName()}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Generated in batches starting ${new Date().toLocaleString()}`);
  doc.saveAndClose();

  const initialState = {
    stack: [{
      id: ROOT_FOLDER_ID,
      prefix: '',
      foldersToken: null, foldersDone: false,
      filesToken: null, filesDone: false,
      hasAnyFiles: null
    }],
    done: false
  };

  props.setProperties({
    'DOC_ID': doc.getId(),
    'STATE': JSON.stringify(initialState)
  });

  Logger.log(`Doc created: ${doc.getUrl()}`);
  processNextFolder();
}

/**
 * Processes as many items as fit in the time budget, checkpointing
 * after every single file/folder, then reschedules itself.
 */
function processNextFolder() {
  deleteTriggersFor(PROCESS_FN_NAME);

  const props = PropertiesService.getScriptProperties();
  const docId = props.getProperty('DOC_ID');
  if (!docId) {
    Logger.log('No state found - run startFileTreeGeneration() first.');
    return;
  }

  const state = JSON.parse(props.getProperty('STATE'));
  if (state.done) {
    Logger.log('Already finished - nothing to do.');
    return;
  }

  const startTime = Date.now();
  const outputLines = [];
  let timeUp = false;

  while (state.stack.length > 0 && !timeUp) {
    const frame = state.stack[state.stack.length - 1];
    const folder = DriveApp.getFolderById(frame.id);

    if (frame.hasAnyFiles === null) {
      // Cheap single check, cached - lets us know if a subfolder item can ever be "last"
      frame.hasAnyFiles = folder.getFiles().hasNext();
    }

    if (!frame.foldersDone) {
      const iter = frame.foldersToken
        ? DriveApp.continueFolderIterator(frame.foldersToken)
        : folder.getFolders();

      if (iter.hasNext()) {
        const sub = iter.next();
        const isLast = !iter.hasNext() && !frame.hasAnyFiles;
        const pointer = isLast ? '└── ' : '├── ';
        outputLines.push(`${frame.prefix}${pointer}📁 ${sub.getName()}`);

        // Checkpoint: resume the PARENT here if we never come back before time runs out
        frame.foldersToken = iter.getContinuationToken();

        // Descend immediately (depth-first)
        state.stack.push({
          id: sub.getId(),
          prefix: frame.prefix + (isLast ? '    ' : '│   '),
          foldersToken: null, foldersDone: false,
          filesToken: null, filesDone: false,
          hasAnyFiles: null
        });
      } else {
        frame.foldersDone = true;
      }
    } else if (!frame.filesDone) {
      const iter = frame.filesToken
        ? DriveApp.continueFileIterator(frame.filesToken)
        : folder.getFiles();

      if (iter.hasNext()) {
        const file = iter.next();
        const isLast = !iter.hasNext();
        const pointer = isLast ? '└── ' : '├── ';
        outputLines.push(`${frame.prefix}${pointer}📄 ${file.getName()}`);
        frame.filesToken = iter.getContinuationToken();
      } else {
        frame.filesDone = true;
      }
    } else {
      // Both phases exhausted for this folder - back to its parent
      state.stack.pop();
    }

    if (Date.now() - startTime > TIME_BUDGET_MS) {
      timeUp = true;
    }
  }

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  outputLines.forEach(line => body.appendParagraph(line));

  if (state.stack.length === 0) {
    state.done = true;
    body.appendParagraph(`✅ Done - full tree generated ${new Date().toLocaleString()}.`).setBold(true);
    doc.saveAndClose();
    props.setProperty('STATE', JSON.stringify(state));
    Logger.log('All done: ' + doc.getUrl());
    deleteTriggersFor(PROCESS_FN_NAME);
    return;
  }

  doc.saveAndClose();
  props.setProperty('STATE', JSON.stringify(state));
  ScriptApp.newTrigger(PROCESS_FN_NAME).timeBased().after(TRIGGER_DELAY_MS).create();
  Logger.log(`Batch done: ${outputLines.length} lines written, stack depth ${state.stack.length}. Continuing...`);
}

/** Clears all saved state and triggers so you can start completely fresh. */
function resetFileTreeGeneration() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  deleteTriggersFor(PROCESS_FN_NAME);
  Logger.log('State cleared. Call startFileTreeGeneration() to start over.');
}

/** Removes any existing triggers pointed at the given function name. */
function deleteTriggersFor(functionName) {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
