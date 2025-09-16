import { Outlet } from "react-router";
import NavBar from "@/components/NavBar";

const RootLayout = () => {
  return (
    <>
      <header>
        <NavBar />
      </header>
      <main>
        <Outlet />
      </main>
    </>
  );
};

export default RootLayout;
