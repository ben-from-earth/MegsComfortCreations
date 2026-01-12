import { MediaType, PreSavedMediaItem } from 'lib/interfaces/globalInterfaces';
import { NextResponse } from 'next/server';

export interface ErrorResponse {
  error: string;
  message: string;
}

export interface SearchErrorResponse extends ErrorResponse {
  failedSearchData: { title: string; author: string } | [];
}

export interface DatabaseSaveEditErrorResponse extends ErrorResponse {
  errors: string[];
  title: string;
}

export class ApiError extends Error {
  status: number;
  error: string;
  constructor(status: number, error: string, message: string) {
    super(message);
    this.name = 'ApiError';
    this.status = status;
    this.error = error;
  }

  format(): NextResponse<ErrorResponse> {
    return NextResponse.json(
      { error: this.error, message: this.message },
      { status: this.status },
    );
  }
}

export class OpenLibraryError extends ApiError {
  failedSearchData: { title: string; author: string };
  constructor(
    status: number,
    error: string,
    message: string,
    failedSearchData: { title: string; author: string },
  ) {
    super(status, error, message);
    this.failedSearchData = failedSearchData;
  }
  format(): NextResponse<ErrorResponse> {
    return NextResponse.json(
      {
        error: this.error,
        message: this.message,
        failedSearchData: this.failedSearchData,
      },
      { status: this.status },
    );
  }
}

export class GoogleSearchError extends ApiError {
  failedSearchData: [];
  constructor(status: number, error: string, message: string) {
    super(status, error, message);
    this.failedSearchData = [];
  }
  format(): NextResponse<ErrorResponse> {
    return NextResponse.json(
      {
        error: this.error,
        message: this.message,
        failedSearchData: this.failedSearchData,
      },
      { status: this.status },
    );
  }
}
