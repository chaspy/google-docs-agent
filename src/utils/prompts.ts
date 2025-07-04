import inquirer from 'inquirer';
import inquirerAutocompletePrompt from 'inquirer-autocomplete-prompt';
import fuzzy from 'fuzzy';
import chalk from 'chalk';
import { User } from '../lib/users';

inquirer.registerPrompt('autocomplete', inquirerAutocompletePrompt);

export class Prompts {
  async selectUser(users: User[]): Promise<string> {
    if (users.length === 0) {
      const { email } = await inquirer.prompt([
        {
          type: 'input',
          name: 'email',
          message: 'Enter the email address of the person to invite:',
          validate: (input: string) => {
            const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
            return emailRegex.test(input) || 'Please enter a valid email address';
          },
        },
      ]);
      return email;
    }

    const { selectedUser } = await inquirer.prompt([
      {
        type: 'autocomplete',
        name: 'selectedUser',
        message: 'Search and select the person to invite:',
        source: async (_: any, input: string) => {
          if (!input) {
            return users.map(user => ({
              name: `${user.name} (${user.email})`,
              value: user.email,
            }));
          }

          const results = fuzzy.filter(input, users, {
            extract: user => `${user.name} ${user.email}`,
          });

          return results.map(result => ({
            name: `${result.original.name} (${result.original.email})`,
            value: result.original.email,
          }));
        },
      },
    ]);

    return selectedUser;
  }

  async getYourName(): Promise<string> {
    const { yourName } = await inquirer.prompt([
      {
        type: 'input',
        name: 'yourName',
        message: 'Enter your name:',
        validate: (input: string) => input.trim().length > 0 || 'Name cannot be empty',
      },
    ]);
    return yourName;
  }

  async confirmCreation(details: {
    title: string;
    editorEmail: string;
    content: string;
  }): Promise<boolean> {
    console.log('\n' + chalk.bold('Document Details:'));
    console.log(chalk.gray('Title:'), details.title);
    console.log(chalk.gray('Editor:'), details.editorEmail);
    console.log(chalk.gray('Initial content preview:'));
    console.log(chalk.dim(details.content.substring(0, 100) + '...'));

    const { confirm } = await inquirer.prompt([
      {
        type: 'confirm',
        name: 'confirm',
        message: 'Create this document?',
        default: true,
      },
    ]);

    return confirm;
  }
}