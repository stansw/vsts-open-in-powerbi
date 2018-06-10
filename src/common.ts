export interface IConfiguration {
    queryId: string;
    contribution: string;
    close?: () => void;
}