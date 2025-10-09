import logo from '@/assets/Comfort.png';
import { NavLink, useNavigate } from 'react-router';
import AccountBoxOutlinedIcon from '@mui/icons-material/AccountBoxOutlined';

const NavBar = () => {
  const navigate = useNavigate();
  const navlinkClass = `w-45 h-4/5 cursor-pointer rounded-md border-3 border-[var(--darkpink)] p-0 flex items-center justify-center font-["Just_Another_Hand"] text-4xl tracking-wider text-black shadow-xl hover:bg-[var(--darkpink)]`;
  return (
    <nav className="p-1.25 relative z-10 flex h-20 items-center gap-4 bg-[var(--lightpink)]">
      <img className="w-18.75 rounded-sm" src={logo} />

      <div className="ml-auto flex h-full flex-row items-center gap-5 pr-5">
        <NavLink
          to="/"
          className={({ isActive }) =>
            `${navlinkClass} ${isActive ? `bg-[var(--darkpink)]` : `bg-[var(--lightpink)]`}`
          }
        >
          Home
        </NavLink>
        <NavLink
          to="ShowDatabase"
          className={({ isActive }) =>
            `${navlinkClass} ${isActive ? `bg-[var(--darkpink)]` : `bg-[var(--lightpink)]`}`
          }
        >
          Show Database
        </NavLink>

        <NavLink
          to="MediaCollector"
          className={({ isActive }) =>
            `${navlinkClass} ${isActive ? `bg-[var(--darkpink)]` : `bg-[var(--lightpink)]`}`
          }
        >
          Media Collector
        </NavLink>
        <AccountBoxOutlinedIcon
          onClick={() => navigate('/Profile', { replace: true })}
          sx={{ fontSize: '70px', p: 0 }}
          className="cursor-pointer text-[var(--darkpink)]"
        />
      </div>
    </nav>
  );
};

export default NavBar;
