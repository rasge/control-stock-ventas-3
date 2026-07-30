import { sql } from "@vercel/postgres";
import { ensureTables } from "../../../lib/db";
import { NextResponse } from "next/server";

export async function GET() {
  try {
    await ensureTables();
    const { rows } = await sql`SELECT * FROM products ORDER BY created_at DESC`;
    return NextResponse.json(rows);
  } catch (e) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}

export async function POST(request) {
  try {
    await ensureTables();
    const body = await request.json();
    const { name, cost, price, stock, stock_minimo, comision_ml } = body;
    if (!name || cost == null || price == null) {
      return NextResponse.json({ error: "Faltan datos" }, { status: 400 });
    }
    const { rows } = await sql`
      INSERT INTO products (name, cost, price, stock, stock_minimo, comision_ml)
      VALUES (${name}, ${cost}, ${price}, ${stock || 0}, ${stock_minimo || 5}, ${comision_ml || 11})
      RETURNING *`;
    return NextResponse.json(rows[0]);
  } catch (e) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}