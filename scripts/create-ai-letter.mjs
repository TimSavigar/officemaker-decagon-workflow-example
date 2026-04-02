import { createDocument, printCreateSummary } from '../src/officemaker-client.mjs';

const documentObject = {
  type: 'document',
  content: {
    children: [
      { type: 'paragraph', children: [{ type: 'text', text: 'Dear Customer,' }] },
      { type: 'paragraph', children: [{ type: 'text', text: 'This document was created from the OfficeMaker Decagon workflow starter.' }] },
      { type: 'paragraph', children: [{ type: 'text', text: 'Kind regards,' }] },
      { type: 'paragraph', children: [{ type: 'text', text: 'Decagon Agent' }] }
    ]
  }
};

const result = await createDocument({ documentType: 'word', fileName: 'decagon-letter', documentObject });
printCreateSummary(result);
