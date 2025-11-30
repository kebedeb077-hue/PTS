import React from 'react';
import './App.css';
import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import Login from './pages/Login';
import Dashboard from './pages/Dashboard';
import FormsList from './pages/FormsList';
import FormDetail from './pages/FormDetail';
import OfficesList from './pages/OfficesList';
import OfficeDetail from './pages/OfficeDetail';
import ChangePassword from './pages/ChangePassword';
import PrivateRoute from './components/PrivateRoute';

function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/login" element={<Login />} />
        <Route element={<PrivateRoute />}>
          <Route path="/" element={<Dashboard />} />
          <Route path="/forms" element={<FormsList />} />
          <Route path="/forms/:id" element={<FormDetail />} />
          <Route path="/offices" element={<OfficesList />} />
          <Route path="/offices/:id" element={<OfficeDetail />} />
          <Route path="/change-password" element={<ChangePassword />} />
        </Route>
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </BrowserRouter>
  );
}

export default App;
