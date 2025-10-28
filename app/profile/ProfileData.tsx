import axios from 'axios';

interface User {
  id: number;
  firstName: string;
  lastName: string;
  email: string;
}

//server location import from .env
const base_URL = process.env.SERVER_BASE_URL;

async function getProfileData(): Promise<User> {
  const res = await axios.get<User>(`${base_URL}/profile`);

  const user = res.data;
  return user;
}

export default async function ProfileData() {
  const user: User = await getProfileData();

  return (
    <div>
      <p>{user.firstName}</p>
      <p>{user.lastName}</p>
      <p>{user.email}</p>
    </div>
  );
}
