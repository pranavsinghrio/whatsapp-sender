import "./globals.css";

export const metadata = {
  title: "WhatsApp Hiring Sender",
  description: "Upload a sheet of HR numbers and message them on WhatsApp.",
};

export const viewport = {
  width: "device-width",
  initialScale: 1,
  maximumScale: 5,
  themeColor: "#25d366",
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
