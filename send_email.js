import { Resend } from 'resend';

const resend = new Resend('re_xxxxxxxxx'); // Replace with your real API key

resend.emails.send({
  from: 'onboarding@resend.dev',
  to: 'hammad@albertsgroup.net',
  subject: 'Hello World',
  html: '<p>Congrats on sending your <strong>first email</strong>!</p>'
});
