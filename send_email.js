import { Resend } from 'resend';

const resend = new Resend('re_UZ6bX6yR_BLL4pdrww5KM1vJnJwQGi2tb');

resend.emails.send({
  from: 'onboarding@resend.dev',
  to: 'hammad@albertsgroup.net',
  subject: 'Hello World',
  html: '<p>Congrats on sending your <strong>first email</strong>!</p>'
});
