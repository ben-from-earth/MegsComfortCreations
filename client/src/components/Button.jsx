import { twMerge } from 'tailwind-merge';

const Button = ({ onClick, label, width, fontSize, additionalStyling }) => {
  const base = `border-3 cursor-pointer rounded-lg border-[var(--darkpink)] bg-[var(--lightpink)] px-2 font-["Just_Another_Hand"] tracking-wider text-black hover:bg-[var(--darkpink)]`;
  return (
    <button
      style={{ width: `${width}px`, fontSize: `${fontSize}px` }}
      className={twMerge(base, additionalStyling)}
      onClick={onClick}
    >
      {label}
    </button>
  );
};

export default Button;
