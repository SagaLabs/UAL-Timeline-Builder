import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'Sagalabs - UAL-Timeline-Builder (UTB)',
  description: 'The tool intended use is to help you in your M365 BEC investigations, or prepare the UAL for import to SIEMs. Made by Christian Henriksen (Guzzy711), SagaLabs',
}

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode
}>) {
  return (
    <html lang="en">
      <head>
      <link rel="icon" href="/favicons/favicon.ico" />
        <link rel="apple-touch-icon" sizes="180x180" href="/favicons/apple-touch-icon.png" />
        <link rel="icon" type="image/png" sizes="32x32" href="/favicons/favicon-32x32.png" />
        <link rel="icon" type="image/png" sizes="16x16" href="/favicons/favicon-16x16.png" />
        <link rel="manifest" href="/favicons/site.webmanifest" />
        <link rel="mask-icon" href="/favicons/safari-pinned-tab.svg" color="#5bbad5" />
      </head>
      <body>{children}</body>
    </html>
  )
}
