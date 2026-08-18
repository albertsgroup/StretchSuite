import { Resend } from 'resend';

if (!process.env.RESEND_API_KEY) {
  throw new Error('RESEND_API_KEY is not set in the environment');
}

const resend = new Resend(process.env.RESEND_API_KEY);

resend.emails.send({
  from: 'Stretch Suite <hello@stretchsuite.com>',
  to: 'hammad@albertsgroup.net',
  subject: 'Hello World',
  html: '<p>Congrats on sending your <strong>first email</strong>!</p>'
});
