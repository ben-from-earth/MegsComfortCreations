'use client';

import React, { useState } from 'react';
import { Formik, Form, useField } from 'formik';
import * as Yup from 'yup';
import { authClient } from 'lib/auth-client';
import { useRouter } from 'next/navigation';

// ---- Formik input component ----

type MyTextInputProps = {
  label: string;
  name: string;
  type?: string;
} & React.InputHTMLAttributes<HTMLInputElement>;

const MyTextInput: React.FC<MyTextInputProps> = ({
  label,
  name,
  type = 'text',
  ...props
}) => {
  const [field, meta] = useField<string>(name);

  const showError = meta.touched && !!meta.error;

  return (
    <>
      <div className="flex w-full items-center gap-4 px-5">
        <label
          className='font-["Just_Another_Hand"] text-4xl'
          htmlFor={props.id || name}
        >
          {label}
        </label>
        <input
          type={type}
          className="ml-auto flex h-10 items-center justify-center rounded bg-white px-2"
          {...props}
          {...field}
          style={{
            border: showError ? '1px solid red' : '1px solid white',
            ...(props.style || {}),
          }}
        />
      </div>
      {showError && <p className="text-lg text-red-600">{meta.error}</p>}
    </>
  );
};

// ---- Form component ----

type LoginValues = {
  email: string;
  password: string;
};

const ItemForm: React.FC<{
  handleSubmit: (values: LoginValues) => Promise<void>;
  loginError: boolean;
}> = ({ handleSubmit, loginError }) => {
  return (
    <Formik<LoginValues>
      initialValues={{ email: '', password: '' }}
      validationSchema={Yup.object({
        email: Yup.string()
          .trim()
          .email('Enter a valid email')
          .required('Email is required')
          .max(254, 'Email must be at most 254 characters'),
        password: Yup.string().required('Password is required'),
      })}
      onSubmit={async (values, { setSubmitting }) => {
        try {
          await handleSubmit(values);
        } finally {
          setSubmitting(false);
        }
      }}
    >
      <Form className="flex flex-col items-center gap-3">
        <MyTextInput label="Email:" name="email" type="email" />
        <MyTextInput label="Password:" name="password" type="password" />
        {loginError && (
          <p className="text-lg text-red-600">Email or Password Invalid</p>
        )}
        <div className="buttonHolder">
          <button
            type="submit"
            className='w-fit cursor-pointer rounded-lg border-3 border-(--darkpink) bg-(--lightpink) p-1 px-2 font-["Just_Another_Hand"] text-4xl tracking-wider text-black hover:bg-(--darkpink)'
          >
            Log In
          </button>
        </div>
      </Form>
    </Formik>
  );
};

const LoginPage: React.FC = () => {
  const router = useRouter();
  const [loginError, setLoginError] = useState(false);

  const handleSubmit = async (values: LoginValues) => {
    setLoginError(false);
    console.log(values);

    const { error } = await authClient.signIn.email({
      email: values.email,
      password: values.password,
      callbackURL: '/profile',
    });

    if (error) {
      setLoginError(true);
      return;
    }

    router.replace('/profile');
    router.refresh();
  };

  return (
    <div className="tracking wider absolute top-1/2 left-1/2 flex h-fit w-fit -translate-x-1/2 -translate-y-1/2 flex-col items-center rounded-md border-3 border-(--darkpink) p-2 shadow-xl">
      <h2 className='mb-4 font-["Just_Another_Hand"] text-4xl'>
        Please Log In!
      </h2>

      <ItemForm handleSubmit={handleSubmit} loginError={loginError} />
    </div>
  );
};

export default LoginPage;
