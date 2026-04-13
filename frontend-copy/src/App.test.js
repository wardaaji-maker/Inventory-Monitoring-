import { render, screen } from '@testing-library/react';
import App from './App';

test('renders Inventory Monitoring title', () => {
  render(<App />);
  const titleElement = screen.getByText(/Inventory Monitoring/i);
  expect(titleElement).toBeInTheDocument();
});
