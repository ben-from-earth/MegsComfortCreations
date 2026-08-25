import { TextField } from '@mui/material';

export default function TextInput({
  onChange,
  onBlur,
  label,
  value,
  rows,
  variant,
  id,
  name,
}: {
  onChange: (
    e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>,
  ) => void;
  onBlur?: (
    e: React.FocusEvent<HTMLInputElement | HTMLTextAreaElement>,
  ) => void;
  label: string;
  value?: string;
  rows?: number;
  variant: 'multiline' | 'normal';
  id?: string;
  name?: string;
  'aria-invalid'?: boolean;
}) {
  return variant === 'multiline' ? (
    <TextField
      className="w-90 rounded-sm bg-white"
      id={id ?? `outlined-multiline-static ${label}`}
      name={name}
      multiline
      label={label}
      slotProps={{
        inputLabel: {
          sx: {
            '&.MuiInputLabel-shrink': {
              backgroundColor: 'white',
              borderRadius: '8px',
              px: '10px',
              color: 'rgb(0,0,0, 0.5)',
              transform: 'translate(6px, -8px) scale(0.75)',
            },
          },
        },
      }}
      rows={rows}
      onChange={onChange}
      onBlur={onBlur}
      value={value}
    />
  ) : (
    <TextField
      className="w-90 rounded-sm bg-white"
      id={id ?? `outlined-static ${label}`}
      name={name}
      label={label}
      onChange={onChange}
      onBlur={onBlur}
      value={value}
      slotProps={{
        inputLabel: {
          sx: {
            '&.MuiInputLabel-shrink': {
              backgroundColor: 'white',
              borderRadius: '8px',
              px: '10px',
              color: 'rgb(0,0,0, 0.5)',
              transform: 'translate(6px, -8px) scale(0.75)',
            },
          },
        },
      }}
    />
  );
}
