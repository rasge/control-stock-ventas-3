export const metadata = {
  title: "Control de Stock y Ventas",
  description: "App para gestionar stock y registrar ventas",
};

import "./globals.css";

export default function RootLayout({ children }) {
  return (
    <html lang="es">
      <body>{children}</body>
    </html>
  );
}