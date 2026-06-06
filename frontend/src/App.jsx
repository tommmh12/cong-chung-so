import { BrowserRouter, Routes, Route } from 'react-router-dom';
import CustomerView from './views/CustomerView';
import AdminView from './views/AdminView';

function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<CustomerView />} />
        <Route path="/admin" element={<AdminView />} />
      </Routes>
    </BrowserRouter>
  );
}

export default App;
