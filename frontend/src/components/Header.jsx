import React from 'react';
import { Link } from 'react-router-dom';
const Navbar = ({ isIndex, isCreate }) => {
  return (
    <nav className="blue darken-3">
      <div className="nav-wrapper">
        <Link to="/" className="brand-logo center hide-on-small-only">
          Вымпел
        </Link>
      </div>
    </nav>
  );
};

export default Navbar;
