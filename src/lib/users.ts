import { google, admin_directory_v1 } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';

export interface User {
  email: string;
  name: string;
}

export class UserService {
  private admin: admin_directory_v1.Admin;

  constructor(auth: OAuth2Client) {
    this.admin = google.admin({ version: 'directory_v1', auth });
  }

  async searchUsers(query: string): Promise<User[]> {
    try {
      const response = await this.admin.users.list({
        customer: 'my_customer',
        query: query ? `email:${query}* name:${query}*` : undefined,
        maxResults: 10,
        orderBy: 'email',
      });

      const users = response.data.users || [];
      
      return users.map(user => ({
        email: user.primaryEmail || '',
        name: user.name?.fullName || user.primaryEmail || '',
      }));
    } catch (error: any) {
      if (error.code === 403) {
        console.error('Note: User search requires Google Workspace admin permissions.');
        console.error('Falling back to manual email input.');
        return [];
      }
      throw error;
    }
  }
}