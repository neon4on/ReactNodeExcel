import Header from './components/Header';
import Home from './pages/Home';
import Page51 from './pages/Form_5_1';
import Page52 from './pages/Form_5_2';
import Page54 from './pages/Form_5_4';
import Page64 from './pages/Form_6_4';
import Page723 from './pages/Form_7_2_3';
import React from 'react';
import { Routes, Route } from 'react-router-dom';
function App() {
  return (
    <div className="wrapper">
      <Header />
      <div className="content">
        <Routes>
          <Route path="*" element={<Home />} />
          <Route path="/51" element={<Page51 />} />
          <Route path="/52" element={<Page52 />} />
          <Route path="/54" element={<Page54 />} />
          <Route path="/64" element={<Page64 />} />
          <Route path="/723" element={<Page723 />} />
        </Routes>
      </div>
    </div>
  );
}

export default App;
