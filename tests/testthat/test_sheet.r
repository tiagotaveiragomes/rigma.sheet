library(rigma.sheet)


test_that("this is working nice", {

  df1 <- as.data.frame(cbind(
    col_a=c(1,2,3),
    col_b=c(3,7,4),
    col_c=c(9,1,3)
  ))
  df2 <- as.data.frame(cbind(
    col_a=c(3,7,4),
    col_b=c(4,5,6),
    col_c=c(NA,NA,NA)
  ))

  stat <- rigma.sheet::Stat$new()
  sheet <- rigma.sheet::Sheet$new(
    name="sheet_name",
    title="Quite nice sheet"
  )
  report <- rigma.sheet::Report$new(
    name="test_report",
    path="~/Desktop",
    gdrive="",
    title="Test Report",
    description="A bery neat report",
    author="Tiago Taveira-Gomes",
    config.filename.ts=FALSE
  )

  sheet$pop.char(
    config=list(
      section.spacer.rows=1
    ),
    styles=list(
      s.any.g.any.a.any.v.std=list(rigma.sheet::styles.color.gray)
    ),
    rules=list(

    ),
    header=list(
      title="Nice",
      labels=list(
        list(
          name="avg"
        ),
        list(
          name="std"
        )
      ),
      attrs=list(
        list(
          name="col_b",
          title="Column B",
          stats=list(
            list(fun=stat$avg),
            list(fun=stat$std)
          )
        ),
        list(
          name="col_c",
          title="Column C",
          stats=list(
            list(fun=stat$avg),
            list(fun=stat$std)
          )
        )
      )
    ),
    sections=list(
      list(
        name="lab_results",
        title="Lab results",
        labels=list(
          list(
            name="avg",
            title="Avg"
          ),
          list(
            name="std",
            title="Std"
          )
        ),
        attrs=list(
          list(
            name="col_b",
            title="Column B",
            stats=list(
              list(fun=stat$avg),
              list(fun=stat$std)
            )
          ),
          list(
            name="col_c",
            title="Column C",
            stats=list(
              list(fun=stat$avg),
              list(fun=stat$std)
            )
          )
        )
      ),
      list(
        name="measurements",
        title="Measurements",
        labels=list(
          list(
            name="avg",
            title="Avg"
          ),
          list(
            name="std",
            title="Std"
          )
        ),
        attrs=list(
          list(
            name="col_b",
            title="Column B",
            stats=list(
              list(fun=stat$avg),
              list(fun=stat$std)
            )
          ),
          list(
            name="col_c",
            title="Column C",
            stats=list(
              list(fun=stat$avg),
              list(fun=stat$std)
            )
          )
        )
      )
    ),
    groups=list(
      list(
        name="group_a",
        title="Group A",
        data=df1
      ),
      list(
        name="group_b",
        title="Group B",
        data=df2
      )
    )
  )
  print("\n")
  print(sheet$data)
  print(digest::digest(sheet$data))
  report$sheet.add(sheet)
  report$save()
})