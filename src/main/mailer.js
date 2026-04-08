import nodemailer from 'nodemailer'

export async function sendEmail({ smtpConfig, to, cc, subject, body, html, attachments }) {
  const transporter = nodemailer.createTransport({
    host: smtpConfig.host,
    port: parseInt(smtpConfig.port),
    secure: smtpConfig.port === '465', // true for 465, false for other ports
    auth: {
      user: smtpConfig.user,
      pass: smtpConfig.pass
    }
  })

  const mailOptions = {
    from: smtpConfig.senderName
      ? `"${smtpConfig.senderName}" <${smtpConfig.user}>`
      : smtpConfig.user,
    to: to,
    cc: cc,
    subject: subject,
    text: body,
    html: html,
    attachments: attachments.map((att) => ({
      filename: att.name,
      path: att.path
    }))
  }

  try {
    const info = await transporter.sendMail(mailOptions)
    return { success: true, messageId: info.messageId }
  } catch (error) {
    console.error('Nodemailer Error:', error)
    return { success: false, error: error.message }
  }
}

export async function verifySmtp(smtpConfig) {
  const transporter = nodemailer.createTransport({
    host: smtpConfig.host,
    port: parseInt(smtpConfig.port),
    secure: smtpConfig.port === '465',
    auth: {
      user: smtpConfig.user,
      pass: smtpConfig.pass
    },
    connectionTimeout: 10000, // 10 seconds
    greetingTimeout: 10000,
    socketTimeout: 10000
  })

  try {
    // Add a hard timeout wrapper to ensure the UI doesn't stay stuck
    const timeoutPromise = new Promise((_, reject) =>
      setTimeout(() => reject(new Error('Connection timed out after 15 seconds')), 15000)
    )

    await Promise.race([transporter.verify(), timeoutPromise])
    return { success: true }
  } catch (error) {
    console.error('SMTP Verification Error:', error)
    return { success: false, error: error.message || 'Unknown verification error' }
  }
}
