import { google, docs_v1, drive_v3 } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import dayjs from 'dayjs';

export class DocsService {
  private docs: docs_v1.Docs;
  private drive: drive_v3.Drive;

  constructor(auth: OAuth2Client) {
    this.docs = google.docs({ version: 'v1', auth });
    this.drive = google.drive({ version: 'v3', auth });
  }

  async createDocument(
    title: string,
    initialContent: string,
    editorEmail: string
  ): Promise<{ documentId: string; documentUrl: string }> {
    const createResponse = await this.docs.documents.create({
      requestBody: {
        title,
      },
    });

    const documentId = createResponse.data.documentId!;

    const requests: docs_v1.Schema$Request[] = [
      {
        insertText: {
          location: {
            index: 1,
          },
          text: initialContent,
        },
      },
    ];

    await this.docs.documents.batchUpdate({
      documentId,
      requestBody: {
        requests,
      },
    });

    await this.drive.permissions.create({
      fileId: documentId,
      requestBody: {
        type: 'user',
        role: 'writer',
        emailAddress: editorEmail,
      },
      fields: 'id',
    });

    const documentUrl = `https://docs.google.com/document/d/${documentId}/edit`;

    return { documentId, documentUrl };
  }

  generateInitialContent(date: string, yourName: string, partnerName: string): string {
    return `# ${date} * ${yourName} * ${partnerName}\n\n## 話したいこと\n\n- \n\n## メモ\n\n`;
  }

  generateDocumentTitle(partnerName: string): string {
    const date = dayjs().format('YYYY-MM-DD');
    return `1on1 - ${date} - ${partnerName}`;
  }
}