import React, { useState, useEffect } from 'react';
import './App.css';

const API_URL = 'http://localhost:3001/api';

function App() {
  const [inventory, setInventory] = useState([]);
  const [categories, setCategories] = useState([]);
  const [activeView, setActiveView] = useState('dashboard'); // dashboard, physical_count, categories

  useEffect(() => {
    const fetchData = async () => {
      try {
        const [inventoryRes, categoriesRes] = await Promise.all([
          fetch(`${API_URL}/inventory`),
          fetch(`${API_URL}/categories`),
        ]);
        const inventoryData = await inventoryRes.json();
        const categoriesData = await categoriesRes.json();
        setInventory(inventoryData);
        setCategories(categoriesData);
      } catch (error) {
        console.error("Error fetching data:", error);
      }
    };
    fetchData();
  }, []);

  const handleCategorySubmit = async (newCategoryName) => {
    try {
      const response = await fetch(`${API_URL}/categories`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: newCategoryName }),
      });
      const newCategory = await response.json();
      setCategories([...categories, newCategory]);
    } catch (error) {
      console.error("Error adding category:", error);
    }
  };

  const handleInventorySubmit = async (newItem) => {
    try {
      const response = await fetch(`${API_URL}/inventory`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(newItem),
      });
      const newInventoryItem = await response.json();
      setInventory([...inventory, newInventoryItem]);
    } catch (error) {
      console.error("Error adding inventory item:", error);
    }
  };

  const handleQuantityUpdate = async (itemId, newQuantity) => {
    try {
      const response = await fetch(`${API_URL}/inventory/${itemId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ quantity: newQuantity }),
      });
      const updatedItem = await response.json();
      setInventory(inventory.map(item =>
        item.id === itemId ? updatedItem : item
      ));
    } catch (error) {
      console.error("Error updating quantity:", error);
    }
  };


  return (
    <div className="App">
      <header className="App-header">
        <h1>Inventory Monitoring</h1>
        <nav>
          <button onClick={() => setActiveView('dashboard')}>Dashboard</button>
          <button onClick={() => setActiveView('physical_count')}>Physical Count</button>
          <button onClick={() => setActiveView('categories')}>Categories</button>
        </nav>
      </header>
      <main>
        {activeView === 'dashboard' && <Dashboard inventory={inventory} />}
        {activeView === 'physical_count' && <PhysicalCount inventory={inventory} onQuantityUpdate={handleQuantityUpdate} />}
        {activeView === 'categories' && <Categories categories={categories} onCategorySubmit={handleCategorySubmit} />}
        <InventoryForm categories={categories} onInventorySubmit={handleInventorySubmit} />
      </main>
    </div>
  );
}

function Dashboard({ inventory }) {
  return (
    <div>
      <h2>Dashboard</h2>
      <table>
        <thead>
          <tr>
            <th>Name</th>
            <th>Category</th>
            <th>Quantity</th>
          </tr>
        </thead>
        <tbody>
          {inventory.map(item => (
            <tr key={item.id}>
              <td>{item.name}</td>
              <td>{item.category}</td>
              <td>{item.quantity}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function PhysicalCount({ inventory, onQuantityUpdate }) {
  const [updates, setUpdates] = useState({});

  const handleUpdateChange = (itemId, quantity) => {
    setUpdates({
      ...updates,
      [itemId]: quantity
    });
  };

  const handleSubmit = (itemId) => {
    const newQuantity = parseInt(updates[itemId], 10);
    if (!isNaN(newQuantity)) {
      onQuantityUpdate(itemId, newQuantity);
      setUpdates({ ...updates, [itemId]: '' }); // Clear input after submit
    }
  };

  return (
    <div>
      <h2>Physical Count</h2>
      <table>
        <thead>
          <tr>
            <th>Name</th>
            <th>Current Quantity</th>
            <th>New Quantity</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {inventory.map(item => (
            <tr key={item.id}>
              <td>{item.name}</td>
              <td>{item.quantity}</td>
              <td>
                <input
                  type="number"
                  value={updates[item.id] || ''}
                  onChange={(e) => handleUpdateChange(item.id, e.target.value)}
                />
              </td>
              <td>
                <button onClick={() => handleSubmit(item.id)}>Update</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function Categories({ categories, onCategorySubmit }) {
  const [newCategory, setNewCategory] = useState('');

  const handleSubmit = (e) => {
    e.preventDefault();
    if (newCategory.trim()) {
      onCategorySubmit(newCategory.trim());
      setNewCategory('');
    }
  };

  return (
    <div>
      <h2>Categories</h2>
      <ul>
        {categories.map(cat => (
          <li key={cat.id}>{cat.name}</li>
        ))}
      </ul>
      <form onSubmit={handleSubmit}>
        <input
          type="text"
          value={newCategory}
          onChange={(e) => setNewCategory(e.target.value)}
          placeholder="New category name"
        />
        <button type="submit">Add Category</button>
      </form>
    </div>
  );
}

function InventoryForm({ categories, onInventorySubmit }) {
  const [name, setName] = useState('');
  const [category, setCategory] = useState('');
  const [quantity, setQuantity] = useState('');

  const handleSubmit = (e) => {
    e.preventDefault();
    if (name.trim() && category && quantity.trim()) {
      onInventorySubmit({
        name: name.trim(),
        category,
        quantity: parseInt(quantity.trim(), 10)
      });
      setName('');
      setCategory('');
      setQuantity('');
    }
  };

  return (
    <div>
      <h2>Add New Inventory Item</h2>
      <form onSubmit={handleSubmit}>
        <input
          type="text"
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder="Item Name"
        />
        <select value={category} onChange={(e) => setCategory(e.target.value)}>
          <option value="" disabled>Select Category</option>
          {categories.map(cat => (
            <option key={cat.id} value={cat.name}>{cat.name}</option>
          ))}
        </select>
        <input
          type="number"
          value={quantity}
          onChange={(e) => setQuantity(e.target.value)}
          placeholder="Quantity"
        />
        <button type="submit">Add Item</button>
      </form>
    </div>
  );
}


export default App;
