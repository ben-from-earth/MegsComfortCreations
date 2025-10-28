// library imports
import { twMerge } from 'tailwind-merge';

// interfaces and types
interface BtnProps {
  onClick?: () => void;
  label: string;
  width: number;
  fontSize: number;
  additionalStyling?: string;
  disabled?: boolean;
}

const Button = ({
  onClick,
  label,
  width,
  fontSize,
  additionalStyling,
  disabled,
}: BtnProps) => {
  const base = `border-3 cursor-pointer rounded-lg border-darkpink bg-lightpink px-2 tracking-wider text-black hover:bg-darkpink shadow-[0px_2px_6px_rgba(0,0,0,0.3)]`;
  return (
    <button
      disabled={disabled || false}
      style={{ width: `${width}px`, fontSize: `${fontSize}px` }}
      className={twMerge(base, additionalStyling)}
      onClick={onClick}
    >
      {label}
    </button>
  );
};

export default Button;
