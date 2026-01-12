import { TextField } from '@mui/material';

export default function TextInput({
  onChange,
  label,
  value,
  rows,
  variant,
}: {
  onChange: (
    e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>,
  ) => void;
  label: string;
  value?: string;
  rows?: number;
  variant: 'multiline' | 'normal';
}) {
  return variant === 'multiline' ? (
    <TextField
      className="w-90 rounded-sm bg-white"
      id={`outlined-multiline-static ${label}`}
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
      value={value}
    />
  ) : (
    <TextField
      className="w-90 rounded-sm bg-white"
      id={`outlined-static ${label}`}
      label={label}
      onChange={onChange}
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
