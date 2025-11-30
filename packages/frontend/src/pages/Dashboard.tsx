import React from 'react';
import { Link } from 'react-router-dom';

export default function Dashboard() {
  return (
    <div>
      <h1>Dashboard</h1>
      <nav>
        <ul>
          <li>
            <Link to="/forms">Forms</Link>
          </li>
          <li>
            <Link to="/offices">Offices</Link>
          </li>
          <li>
            <Link to="/change-password">Change Password</Link>
          </li>
        </ul>
      </nav>
    </div>
  );
}
