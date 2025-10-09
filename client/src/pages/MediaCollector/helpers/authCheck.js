import axios from 'axios';
import { redirect } from 'react-router';

//server location import from .env
const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

const requireAuth = async () => {
  const res = await axios.get(`${serverDomain}/auth/me`, {
    withCredentials: true,
    validateStatus: (status) => status < 500,
  });

  if (res.status === 200) {
    return res.data.user;
  }
  throw redirect(`/login`);
};

export default requireAuth;
