// library imports
import { twMerge } from 'tailwind-merge';
import React from 'react';

import UnfoldMoreIcon from '@mui/icons-material/UnfoldMore';
import CloseIcon from '@mui/icons-material/Close';

// interfaces and types
interface BtnProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  onClick?: React.MouseEventHandler<HTMLButtonElement>;
  label?: string;
  width?: number;
  fontSize?: number;
  className?: string;
  variant: 'primary' | 'popover' | 'close';
  disabled?: boolean;
}

const buttonBaseClasses = {
  primary: `border-3 cursor-pointer rounded-lg border-darkpink bg-lightpink px-2 tracking-wider text-black hover:bg-darkpink shadow-[0px_2px_6px_rgba(0,0,0,0.3)]`,
  popover: `border-3 cursor-pointer rounded-lg border-darkpink bg-lightpink px-2 tracking-wider text-black hover:bg-darkpink shadow-[0px_2px_6px_rgba(0,0,0,0.3)] flex items-center justify-between`,
  close: `cursor-pointer hover:bg-darkpink flex items-center justify-center`,
};

const Button = ({
  onClick,
  label,
  width,
  fontSize,
  disabled,
  variant,
  className,
  type = 'button',
  ...rest
}: BtnProps) => {
  return (
    <button
      type={type}
      disabled={disabled || false}
      style={{ width: `${width}px`, fontSize: `${fontSize}px` }}
      className={twMerge(buttonBaseClasses[variant], className)}
      onClick={onClick}
      {...rest}
    >
      {label && <span>{label}</span>}{' '}
      {variant === 'popover' && <UnfoldMoreIcon className="opacity-50" />}
      {variant === 'close' && <CloseIcon fontSize="small" />}
    </button>
  );
};

export default Button;
