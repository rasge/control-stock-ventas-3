import { sql } from "@vercel/postgres";
import { ensureTables } from "../../../lib/db";
import { NextResponse } from "next/server";

export async function GET() {
  try {
    await ensureTables();
    const { rows } = await sql`
      SELECT s.*, p.name as product_name
      FROM sales s
      JOIN products p ON s.product_id = p.id
      ORDER BY s.sale_date DESC`;
    return NextResponse.json(rows);
  } catch (e) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}

export async function POST(request) {
  try {
    await ensureTables();
    const body = await request.json();
    const { product_id, quantity, sale_date } = body;
    if (!product_id || !quantity) {
      return NextResponse.json({ error: "Faltan datos" }, { status: 400 });
    }
    const { rows: productRows } = await sql`SELECT * FROM products WHERE id = ${product_id}`;
    if (productRows.length === 0) {
      return NextResponse.json({ error: "Producto no encontrado" }, { status: 404 });
    }
    const product = productRows[0];
    const profit = (Number(product.price) - Number(product.cost)) * Number(quantity);
    const { rows } = await sql`
      INSERT INTO sales (product_id, quantity, sale_date, profit)
      VALUES (${product_id}, ${quantity}, ${sale_date}, ${profit})
      RETURNING *`;
    await sql`UPDATE products SET stock = stock - ${quantity} WHERE id = ${product_id}`;
    return NextResponse.json(rows[0]);
  } catch (e) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}