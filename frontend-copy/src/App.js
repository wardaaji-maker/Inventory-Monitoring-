import React, { useState, useEffect } from 'react';
import './App.css';

// Mock data for development
const mockInventory = [
  { id: 1, name: 'Laptop', category: 'Electronics', quantity: 15 },
  { id: 2, name: 'Keyboard', category: 'Electronics', quantity: 50 },
  { id: 3, name: 'Mouse', category: 'Electronics', quantity: 75 },
  { id: 4, name: 'Chair', category: 'Furniture', quantity: 20 },
];

const mockCategories = [
  { id: 1, name: 'Electronics' },
  { id: 2, name: 'Furniture' },
  { id: 3, name: 'Office Supplies' },
];


function App() {
  const [inventory, setInventory] = useState([]);
  const [categories, setCategories] = useState([]);
  const [activeView, setActiveView] = useState('dashboard'); // dashboard, physical_count, categories

  useEffect(() => {
    // In a real app, you'd fetch this data from an API.
    // For now, we'll use mock data.
    setInventory(mockInventory);
    setCategories(mockCategories);
  }, []);

  const handleCategorySubmit = (newCategoryName) => {
    const newCategory = {
      id: Date.now(),
      name: newCategoryName,
    };
    setCategories([...categories, newCategory]);
  };

  const handleInventorySubmit = (newItem) => {
    const newInventoryItem = {
      id: Date.now(),
      ...newItem,
    };
    setInventory([...inventory, newInventoryItem]);
  };

  const handleQuantityUpdate = (itemId, newQuantity) => {
    setInventory(inventory.map(item =>
      item.id === itemId ? { ...item, quantity: newQuantity } : item
    ));
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
