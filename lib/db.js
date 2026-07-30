import { sql } from "@vercel/postgres";

let tablesInitialized = false;
let initPromise = null;

export async function ensureTables() {
  if (tablesInitialized) return;
  if (initPromise) return initPromise;

  initPromise = (async () => {
    try {
      await sql`CREATE TABLE IF NOT EXISTS products (
        id SERIAL PRIMARY KEY,
        name TEXT NOT NULL,
        cost NUMERIC NOT NULL,
        price NUMERIC NOT NULL,
        stock INTEGER NOT NULL DEFAULT 0,
        stock_minimo INTEGER DEFAULT 5,
        comision_ml DECIMAL(5,2) DEFAULT 11.00,
        created_at TIMESTAMP DEFAULT now()
      )`;

      await sql`CREATE TABLE IF NOT EXISTS sales (
        id SERIAL PRIMARY KEY,
        product_id INTEGER REFERENCES products(id) ON DELETE CASCADE,
        quantity INTEGER NOT NULL,
        sale_date DATE NOT NULL,
        profit NUMERIC NOT NULL,
        created_at TIMESTAMP DEFAULT now()
      )`;

      tablesInitialized = true;
    } catch (e) {
      console.error("Error creating tables:", e);
      throw e;
    }
  })();

  return initPromise;
}