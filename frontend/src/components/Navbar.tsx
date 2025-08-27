import React from "react";
const Navbar: React.FC = () => (
  <nav className="fixed top-0 left-0 w-full z-50 bg-[#26225b] shadow flex items-center" style={{ height: "64px" }}>
    <div className="max-w-7xl w-full flex items-center px-8 h-full">
      <img src="/prolifics-logo-white.webp" alt="Prolifics" className="h-10 w-auto block" style={{ objectFit: "contain" }}/>
    </div>
  </nav>
);
export default Navbar;
