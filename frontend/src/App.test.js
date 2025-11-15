import { render, screen } from '@testing-library/react';
import App from './App';

test('renders inventory monitoring header', () => {
  render(<App />);
  const headerElement = screen.getByText(/Inventory Monitoring/i);
  expect(headerElement).toBeInTheDocument();
});
