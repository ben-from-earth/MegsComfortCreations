import { NavLink } from 'react-router';

const ErrorBoundary = () => {
  const navlinkClass = `w-45 h-4/5 cursor-pointer rounded-md border-3 border-[var(--darkpink)] p-0 flex items-center justify-center font-["Just_Another_Hand"] text-4xl tracking-wider text-black shadow-xl hover:bg-[var(--darkpink)]`;
  return (
    <div
      className={`border-3 tracking wider absolute left-1/2 top-1/2 flex h-fit w-fit -translate-x-1/2 -translate-y-1/2 flex-col items-center gap-2 rounded-md border-[var(--darkpink)] p-2 font-["Just_Another_Hand"] text-4xl shadow-xl`}
    >
      <p>Sorry! The page you are looking for doesnt exist</p>
      <NavLink
        to="/"
        className={({ isActive }) =>
          `${navlinkClass} ${isActive ? `bg-[var(--darkpink)]` : `bg-[var(--lightpink)]`}`
        }
      >
        Home
      </NavLink>
    </div>
  );
};

export default ErrorBoundary;
