% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/gks.R
\name{Add_OKATO}
\alias{Add_OKATO}
\title{Добавляет коды ОКАТО субъектов РФ к фрейму, содержащему панельные данные, необходим файл subjects.csv, который содержит коды ОКАТО}
\usage{
Add_OKATO(dann_frame_main, id = 1)
}
\arguments{
\item{dann_frame_main}{- фрейм данных}

\item{id}{- индекс столбца, содержащего названия регионов, по умолчанию 1}
}
\value{
исходный фрейм данных со столбцом OKATO
}
\description{
Добавляет коды ОКАТО субъектов РФ к фрейму, содержащему панельные данные, необходим файл subjects.csv, который содержит коды ОКАТО
}
\examples{
Add_OKATO(dataGKS)
}
