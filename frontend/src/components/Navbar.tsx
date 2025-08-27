import React from "react";

const Navbar: React.FC = () => (
  <nav
    className="fixed top-0 left-0 w-full z-50 bg-[#26225b] shadow"
    style={{ minHeight: "64px" }}
  >
    <div className="max-w-7xl mx-auto flex items-center px-8 py-3">
      <img
        src="/prolifics-logo-white.webp" // Logo in public/
        alt="Prolifics"
        className="h-10 w-auto"
        style={{ objectFit: "contain" }}
      />
    </div>
  </nav>
);

export default Navbar;
