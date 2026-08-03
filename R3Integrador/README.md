# R3 Integrador

O projeto principal da solução é o **R3Integrador.Web**, a interface web para importação e acompanhamento de atualizações de tabelas de preço.

```bash
dotnet run --project src/R3Integrador.Web/R3Integrador.Web.csproj
```

Os projetos `Application`, `Infrastructure` e `Core` dão suporte à aplicação web. O `R3Integrador.Console` permanece na solução apenas para compatibilidade com o fluxo legado.
