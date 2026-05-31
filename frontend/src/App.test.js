import { render, screen } from '@testing-library/react';
import App from './App';

test('renders inventory monitoring heading', () => {
  render(<App />);
  const headingElement = screen.getByText(/Inventory Monitoring/i);
  expect(headingElement).toBeInTheDocument();
});
