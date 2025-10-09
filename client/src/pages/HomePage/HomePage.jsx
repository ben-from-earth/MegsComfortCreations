import { NavLink } from 'react-router';

const HomePage = () => {
  const navlinkClass = ` bg-[var(--lightpink)] w-45 h-4/5 cursor-pointer rounded-md border-3 border-[var(--darkpink)] p-2 flex items-center justify-center font-["Just_Another_Hand"] text-4xl tracking-wider text-black shadow-xl hover:bg-[var(--darkpink)]`;

  return (
    <div className="flex flex-col items-center">
      <h1 className='m-5 text-center font-["Just_Another_Hand"] text-7xl tracking-wider'>
        Welcome to Meg's Comfort Creations!
      </h1>

      <div className='flex gap-2 font-["Just_Another_Hand"] text-4xl tracking-wider'>
        <NavLink to="/login" className={navlinkClass}>
          Login
        </NavLink>
        <NavLink to="/signup" className={navlinkClass}>
          Sign Up
        </NavLink>
      </div>
    </div>
  );
};

export default HomePage;
