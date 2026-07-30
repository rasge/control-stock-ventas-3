"use client";

import { useEffect, useState, useCallback } from "react";
import * as XLSX from "xlsx";

const money = (n) =>
  new Intl.NumberFormat("es-AR", { style: "currency", currency: "ARS" }).format(Number(n) || 0);

function normalizeRow(r) {
  const keys = Object.keys(r);
  const find = (opts) => {
    const k = keys.find((k) => opts.includes(k.toLowerCase().trim()));
    return k !== undefined ? r[k] : undefined;
  };
  return {
    name: find(["nombre", "producto", "name", "título", "titulo"]),
    cost: Number(find(["costo", "cost", "precio de costo"])) || 0,
    price: Number(find(["precio", "price", "precio de venta", "precio venta"])) || 0,
    stock: Number(find(["stock", "cantidad", "qty"])) || 0,
    stock_minimo: Number(find(["stock minimo", "stock_minimo", "minimo"])) || 5,
    comision_ml: Number(find(["comision ml", "comision_ml", "comisión"])) || 11,
  };
}

function calcularGananciaNeta(product) {
  const comision = (Number(product.comision_ml) || 11) / 100;
  const precioNeto = Number(product.price) * (1 - comision);
  const gananciaNeta = precioNeto - Number(product.cost);
  const margenNeto = Number(product.price) > 0 ? ((gananciaNeta / Number(product.price)) * 100).toFixed(1) : 0;
  return {
    precioNeto: precioNeto.toFixed(2),
    gananciaNeta: gananciaNeta.toFixed(2),
    margenNeto: margenNeto
  };
}

export default function Home() {
  const [tab, setTab] = useState("productos");
  const [products, setProducts] = useState([]);
  const [sales, setSales] = useState([]);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);
  const [form, setForm] = useState({ name: "", cost: "", price: "", stock: "", stock_minimo: 5, comision_ml: 11 });
  const [saleForm, setSaleForm] = useState({ product_id: "", quantity: "", sale_date: "" });
  const [importing, setImporting] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [editingValues, setEditingValues] = useState({});

  const loadAll = useCallback(async () => {
    try {
      setError("");
      const [pRes, sRes] = await Promise.all([fetch("/api/products"), fetch("/api/sales")]);
      const pJson = await pRes.json();
      const sJson = await sRes.json();
      if (!pRes.ok) throw new Error(pJson.error || "Error cargando productos");
      if (!sRes.ok) throw new Error(sJson.error || "Error cargando ventas");
      setProducts(pJson);
      setSales(sJson);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadAll();
  }, [loadAll]);

  async function addProduct(e) {
    e.preventDefault();
    setError("");
    try {
      const res = await fetch("/api/products", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: form.name,
          cost: Number(form.cost),
          price: Number(form.price),
          stock: Number(form.stock) || 0,
          stock_minimo: Number(form.stock_minimo) || 5,
          comision_ml: Number(form.comision_ml) || 11,
        }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "Error al crear producto");
      setForm({ name: "", cost: "", price: "", stock: "", stock_minimo: 5, comision_ml: 11 });
      loadAll();
    } catch (e) {
      setError(e.message);
    }
  }

  async function deleteProduct(id) {
    if (!confirm("¿Eliminar este producto?")) return;
    try {
      const res = await fetch(`/api/products/${id}`, { method: "DELETE" });
      if (!res.ok) throw new Error("Error al eliminar");
      loadAll();
    } catch (e) {
      setError(e.message);
    }
  }

  async function updateProduct(id) {
    if (!editingId) return;
    setError("");
    try {
      const product = products.find(p => p.id === id);
      const res = await fetch(`/api/products/${id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: editingValues[id]?.name || product.name,
          cost: Number(editingValues[id]?.cost) || Number(product.cost),
          price: Number(editingValues[id]?.price) || Number(product.price),
          stock: Number(editingValues[id]?.stock) !== undefined ? Number(editingValues[id].stock) : Number(product.stock),
          stock_minimo: Number(editingValues[id]?.stock_minimo) || Number(product.stock_minimo) || 5,
          comision_ml: Number(editingValues[id]?.comision_ml) || Number(product.comision_ml) || 11,
        }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "Error al actualizar");
      setEditingId(null);
      setEditingValues({});
      loadAll();
    } catch (e) {
      setError(e.message);
    }
  }

  function handleEditChange(id, field, value) {
    setEditingValues(prev => ({
      ...prev,
      [id]: {
        ...prev[id],
        [field]: value
      }
    }));
  }

  async function registerSale(e) {
    e.preventDefault();
    setError("");
    try {
      const res = await fetch("/api/sales", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          product_id: Number(saleForm.product_id),
          quantity: Number(saleForm.quantity),
          sale_date: saleForm.sale_date || new Date().toISOString().slice(0, 10),
        }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "Error al registrar venta");
      setSaleForm({ product_id: "", quantity: "", sale_date: "" });
      loadAll();
    } catch (e) {
      setError(e.message);
    }
  }

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setImporting(true);
    setError("");
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = evt.target.result;
        const wb = XLSX.read(data, { type: "binary" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);
        const rows = json.map(normalizeRow).filter((r) => r.name);
        if (rows.length === 0) {
          throw new Error("No se encontraron filas válidas");
        }
        const res = await fetch("/api/products/import", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ rows }),
        });
        const resJson = await res.json();
        if (!res.ok) throw new Error(resJson.error || "Error al importar");
        await loadAll();
        alert(`Importados: ${resJson.inserted}. Omitidos: ${resJson.skipped}.`);
      } catch (err) {
        setError(err.message);
      } finally {
        setImporting(false);
        e.target.value = "";
      }
    };
    reader.readAsBinaryString(file);
  }

  const productosBajoStock = products.filter(p => Number(p.stock) <= (Number(p.stock_minimo) || 5));
  const totalStock = products.reduce((a, p) => a + Number(p.stock), 0);
  const totalProfitPotential = products.reduce((a, p) => {
    const g = calcularGananciaNeta(p);
    return a + (Number(g.gananciaNeta) * Number(p.stock));
  }, 0);
  const totalProfitRealized = sales.reduce((a, s) => a + Number(s.profit), 0);

  return (
    <div className="container">
      <h1>Control de Stock y Ventas</h1>
      <p className="subtitle">Cargá tus productos, importá desde Excel y registrá ventas con ganancia automática.</p>

      {error && <div className="error-banner">{error}</div>}

      {productosBajoStock.length > 0 && (
        <div style={{
          backgroundColor: "#fee2e2",
          border: "2px solid #dc2626",
          borderRadius: "8px",
          padding: "16px",
          marginBottom: "20px"
        }}>
          <h3 style={{ margin: "0 0 12px 0", color: "#991b1b" }}>
            🔴 REPONER STOCK - {productosBajoStock.length} producto(s)
          </h3>
          {productosBajoStock.map(p => (
            <div key={p.id} style={{
              padding: "8px",
              marginBottom: "8px",
              backgroundColor: "#fecaca",
              borderRadius: "4px",
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center"
            }}>
              <span>
                <strong>{p.name}</strong> - Stock: {p.stock} unidades (mínimo: {p.stock_minimo || 5})
              </span>
              <a 
                href="https://temu.com" 
                target="_blank" 
                rel="noopener noreferrer"
                style={{
                  padding: "6px 12px",
                  backgroundColor: "#dc2626",
                  color: "white",
                  borderRadius: "4px",
                  textDecoration: "none",
                  fontSize: "12px"
                }}
              >
                Ir a Temu
              </a>
            </div>
          ))}
        </div>
      )}

      <div className="summary-grid">
        <div className="summary-box">
          <div className="label">Productos</div>
          <div className="value">{products.length}</div>
        </div>
        <div className="summary-box">
          <div className="label">Unidades en stock</div>
          <div className="value">{totalStock}</div>
        </div>
        <div className="summary-box">
          <div className="label">Ganancia potencial (neta)</div>
          <div className="value">{money(totalProfitPotential)}</div>
        </div>
        <div className="summary-box">
          <div className="label">Ganancia acumulada</div>
          <div className="value">{money(totalProfitRealized)}</div>
        </div>
      </div>

      <div className="tabs">
        <button className={`tab ${tab === "productos" ? "active" : ""}`} onClick={() => setTab("productos")}>
          Productos
        </button>
        <button className={`tab ${tab === "ventas" ? "active" : ""}`} onClick={() => setTab("ventas")}>
          Ventas
        </button>
      </div>

      {tab === "productos" && (
        <>
          <div className="card">
            <h2>Nuevo producto</h2>
            <form className="inline" onSubmit={addProduct}>
              <div className="field">
                <label>Nombre</label>
                <input required value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
              </div>
              <div className="field">
                <label>Costo</label>
                <input required type="number" step="0.01" value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} />
              </div>
              <div className="field">
                <label>Precio de venta</label>
                <input required type="number" step="0.01" value={form.price} onChange={(e) => setForm({ ...form, price: e.target.value })} />
              </div>
              <div className="field">
                <label>Stock inicial</label>
                <input type="number" value={form.stock} onChange={(e) => setForm({ ...form, stock: e.target.value })} />
              </div>
              <div className="field">
                <label>Stock mínimo</label>
                <input type="number" value={form.stock_minimo} onChange={(e) => setForm({ ...form, stock_minimo: e.target.value })} />
              </div>
              <div className="field">
                <label>Comisión ML (%)</label>
                <input type="number" step="0.1" value={form.comision_ml} onChange={(e) => setForm({ ...form, comision_ml: e.target.value })} />
              </div>
              <button className="primary" type="submit">Agregar</button>
              <label className="file-label">
                {importing ? "Importando…" : "Importar Excel/CSV"}
                <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} disabled={importing} />
              </label>
            </form>
          </div>

          <div className="card">
            <h2>Productos</h2>
            {loading ? (
              <div className="empty">Cargando…</div>
            ) : products.length === 0 ? (
              <div className="empty">Todavía no cargaste productos.</div>
            ) : (
              <table>
                <thead>
                  <tr>
                    <th>Nombre</th>
                    <th>Costo</th>
                    <th>Precio</th>
                    <th>Ganancia Neta</th>
                    <th>Margen</th>
                    <th>Comisión ML</th>
                    <th>Stock Mín.</th>
                    <th>Stock</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {products.map((p) => {
                    const g = calcularGananciaNeta(p);
                    return (
                      <tr key={p.id} style={Number(p.stock) <= (Number(p.stock_minimo) || 5) ? { backgroundColor: "#fef2f2" } : {}}>
                        <td>{editingId === p.id ? <input type="text" value={editingValues[p.id]?.name || p.name} onChange={(e) => handleEditChange(p.id, "name", e.target.value)} /> : p.name}</td>
                        <td>{editingId === p.id ? <input type="number" step="0.01" value={editingValues[p.id]?.cost !== undefined ? editingValues[p.id].cost : p.cost} onChange={(e) => handleEditChange(p.id, "cost", e.target.value)} /> : money(p.cost)}</td>
                        <td>{editingId === p.id ? <input type="number" step="0.01" value={editingValues[p.id]?.price !== undefined ? editingValues[p.id].price : p.price} onChange={(e) => handleEditChange(p.id, "price", e.target.value)} /> : money(p.price)}</td>
                        <td className={Number(g.gananciaNeta) >= 0 ? "profit-positive" : "profit-negative"}>{money(g.gananciaNeta)}</td>
                        <td>{g.margenNeto}%</td>
                        <td>{editingId === p.id ? <input type="number" step="0.1" value={editingValues[p.id]?.comision_ml !== undefined ? editingValues[p.id].comision_ml : (p.comision_ml || 11)} onChange={(e) => handleEditChange(p.id, "comision_ml", e.target.value)} style={{ width: "60px" }} /> : <>{p.comision_ml || 11}%</>}</td>
                        <td>{editingId === p.id ? <input type="number" value={editingValues[p.id]?.stock_minimo !== undefined ? editingValues[p.id].stock_minimo : (p.stock_minimo || 5)} onChange={(e) => handleEditChange(p.id, "stock_minimo", e.target.value)} style={{ width: "50px" }} /> : p.stock_minimo || 5}</td>
                        <td style={Number(p.stock) <= (Number(p.stock_minimo) || 5) ? { color: "#dc2626", fontWeight: "bold" } : {}}>{p.stock}</td>
                        <td>{editingId === p.id ? <><button className="success" onClick={() => updateProduct(p.id)}>Guardar</button><button className="secondary" onClick={() => { setEditingId(null); setEditingValues({}); }}>Cancelar</button></> : <><button className="secondary" onClick={() => { setEditingId(p.id); setEditingValues({}); }}>Editar</button><button className="danger" onClick={() => deleteProduct(p.id)}>Eliminar</button></>}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>
        </>
      )}

      {tab === "ventas" && (
        <>
          <div className="card">
            <h2>Registrar venta</h2>
            <form className="inline" onSubmit={registerSale}>
              <div className="field">
                <label>Producto</label>
                <select required value={saleForm.product_id} onChange={(e) => setSaleForm({ ...saleForm, product_id: e.target.value })}>
                  <option value="">Seleccionar…</option>
                  {products.map((p) => (
                    <option key={p.id} value={p.id}>
                      {p.name} (stock: {p.stock})
                    </option>
                  ))}
                </select>
              </div>
              <div className="field">
                <label>Cantidad</label>
                <input required type="number" min="1" value={saleForm.quantity} onChange={(e) => setSaleForm({ ...saleForm, quantity: e.target.value })} />
              </div>
              <div className="field">
                <label>Fecha</label>
                <input type="date" value={saleForm.sale_date} onChange={(e) => setSaleForm({ ...saleForm, sale_date: e.target.value })} />
              </div>
              <button className="primary" type="submit">Registrar venta</button>
            </form>
          </div>

          <div className="card">
            <h2>Historial de ventas</h2>
            {loading ? (
              <div className="empty">Cargando…</div>
            ) : sales.length === 0 ? (
              <div className="empty">Todavía no registraste ventas.</div>
            ) : (
              <table>
                <thead>
                  <tr>
                    <th>Fecha</th>
                    <th>Producto</th>
                    <th>Cantidad</th>
                    <th>Ganancia</th>
                  </tr>
                </thead>
                <tbody>
                  {sales.map((s) => (
                    <tr key={s.id}>
                      <td>{s.sale_date?.slice(0, 10)}</td>
                      <td>{s.product_name}</td>
                      <td>{s.quantity}</td>
                      <td className={Number(s.profit) >= 0 ? "profit-positive" : "profit-negative"}>{money(s.profit)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </>
      )}
    </div>
  );
}