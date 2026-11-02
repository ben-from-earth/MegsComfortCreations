import Link from 'next/link';

export default function ProfileButton() {
  return (
    <Link
      href="/profile"
      aria-label="Profile"
      className="ml-auto flex size-11 items-center justify-center rounded-full bg-transparent text-[#2b2826]"
    >
      <svg
        width="26"
        height="26"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden="true"
      >
        <circle cx="12" cy="8" r="4" />
        <path d="M4 21c1.5-4 4.5-6 8-6s6.5 2 8 6" />
      </svg>
    </Link>
  );
}
