'use client';
import Link from 'next/link';
import AccountBoxOutlinedIcon from '@mui/icons-material/AccountBoxOutlined';

export default function ProfileButton() {
  return (
    <Link href={'/profile'}>
      <AccountBoxOutlinedIcon
        sx={{ fontSize: '70px', p: 0 }}
        className="text-darkpink cursor-pointer"
      />
    </Link>
  );
}
