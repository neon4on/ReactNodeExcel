import React from 'react';
import { Link } from 'react-router-dom';
const Home = () => {
  return (
    <>
      <div className="h2Home">
        <h2>Vimpel Forms</h2>
      </div>
      <ul className="ulHome">
        <li>
          <Link to="/51">Форма 5.1</Link>
        </li>
        <li>
          <Link to="/52">Форма 5.2</Link>
        </li>
        <li>
          <Link to="/54">Форма 5.4</Link>
        </li>
        <li>
          <Link to="/64">Форма 6.4</Link>
        </li>
        <li>
          <Link to="/723">Форма 7.2.3</Link>
        </li>
      </ul>
    </>
  );
};

export default Home;
