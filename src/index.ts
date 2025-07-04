#!/usr/bin/env node

import { Command } from 'commander';
import chalk from 'chalk';
import dayjs from 'dayjs';
import { GoogleAuth } from './lib/auth';
import { DocsService } from './lib/docs';
import { UserService, User } from './lib/users';
import { Prompts } from './utils/prompts';

const program = new Command();

program
  .name('google-docs-agent')
  .description('Automated Google Docs creation tool for 1on1 meetings')
  .version('1.0.0');

program
  .command('create')
  .description('Create a new 1on1 document')
  .option('-n, --name <name>', 'Your name')
  .option('-e, --email <email>', 'Email of the person to invite')
  .action(async (options) => {
    try {
      console.log(chalk.bold('🚀 Google Docs Agent - 1on1 Document Creator\n'));

      console.log(chalk.yellow('Authenticating with Google...'));
      const auth = new GoogleAuth();
      const client = await auth.authorize();
      console.log(chalk.green('✓ Authentication successful\n'));

      const docsService = new DocsService(client);
      const userService = new UserService(client);
      const prompts = new Prompts();

      const yourName = options.name || await prompts.getYourName();

      let editorEmail: string;
      if (options.email) {
        editorEmail = options.email;
      } else {
        console.log(chalk.yellow('Searching for users...'));
        let users: User[] = [];
        try {
          users = await userService.searchUsers('');
        } catch (error) {
          console.log(chalk.yellow('Could not search users (admin permissions required).'));
        }
        
        editorEmail = await prompts.selectUser(users);
      }

      const partnerName = editorEmail.split('@')[0];
      const date = dayjs().format('YYYY-MM-DD');
      const title = docsService.generateDocumentTitle(partnerName);
      const content = docsService.generateInitialContent(date, yourName, partnerName);

      const shouldCreate = await prompts.confirmCreation({
        title,
        editorEmail,
        content,
      });

      if (!shouldCreate) {
        console.log(chalk.yellow('\nDocument creation cancelled.'));
        return;
      }

      console.log(chalk.yellow('\nCreating document...'));
      const { documentId, documentUrl } = await docsService.createDocument(
        title,
        content,
        editorEmail
      );

      console.log(chalk.green('\n✅ Document created successfully!'));
      console.log(chalk.bold('\nDocument URL:'), chalk.cyan(documentUrl));
      console.log(chalk.gray(`Document ID: ${documentId}`));
      console.log(chalk.gray(`Shared with: ${editorEmail}`));

    } catch (error: any) {
      console.error(chalk.red('\n❌ Error:'), error.message);
      if (error.message.includes('credentials')) {
        console.log(chalk.yellow('\nPlease ensure you have a credentials.json file in the current directory.'));
        console.log(chalk.yellow('You can get this from the Google Cloud Console:'));
        console.log(chalk.cyan('https://console.cloud.google.com/apis/credentials'));
      }
      process.exit(1);
    }
  });

program
  .command('setup')
  .description('Setup Google API credentials')
  .action(() => {
    console.log(chalk.bold('Google Docs Agent Setup\n'));
    console.log('To use this tool, you need to:');
    console.log('\n1. Go to the Google Cloud Console:');
    console.log(chalk.cyan('   https://console.cloud.google.com'));
    console.log('\n2. Create a new project or select an existing one');
    console.log('\n3. Enable the following APIs:');
    console.log('   - Google Docs API');
    console.log('   - Google Drive API');
    console.log('   - Admin SDK API (optional, for user search)');
    console.log('\n4. Create OAuth 2.0 credentials:');
    console.log('   - Go to APIs & Services > Credentials');
    console.log('   - Create credentials > OAuth client ID');
    console.log('   - Application type: Desktop app');
    console.log('   - Download the credentials');
    console.log('\n5. Save the credentials as "credentials.json" in this directory');
    console.log('\n6. Run "google-docs-agent create" to start creating documents!');
  });

program.parse();

if (!process.argv.slice(2).length) {
  program.outputHelp();
}