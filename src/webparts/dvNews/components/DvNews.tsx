import * as React from "react";
import styles from "./DvNews.module.scss";
import type { IDvNewsProps } from "./IDvNewsProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { timeAgo } from "../../../service/helper";
import moment from "moment";
import { Swiper, SwiperSlide } from "swiper/react";
import { Navigation, Pagination } from "swiper/modules";
import "swiper/css";
import "swiper/css/navigation";
import "swiper/css/pagination";
import { Dropdown } from "@fluentui/react/lib/Dropdown";
import { Checkbox, Spinner, SpinnerSize, TextField } from "@fluentui/react";
import { graphfi, SPFx as graphSPFx } from "@pnp/graph";
import "@pnp/graph/groups";
import "@pnp/graph/members";
const DvNews: React.FC<IDvNewsProps> = (props) => {
  const [newsItems, setNewsItems] = React.useState<any[]>([]);
  const [deptOptions, setDeptOptions] = React.useState<any[]>([]);
  const [redirecturl, setRedirectUrl] = React.useState<string>("");
  const [isAddDialogOpen, setIsAddDialogOpen] = React.useState(false);
  const [isLoadingNews, setIsLoadingNews] = React.useState(false);
  const [isCreatingItem, setIsCreatingItem] = React.useState(false);
  const [isLoadingMoreNews, setIsLoadingMoreNews] = React.useState(false);
  const [createdItemSuccess, setCreatedItemSuccess] = React.useState(false);
  const [newTitle, setNewTitle] = React.useState("");
  const [selectedDepartments, setSelectedDepartments] = React.useState<
    string[]
  >([]);
  const [selectedSiteTypes, setSelectedSiteTypes] = React.useState<string[]>(
    [],
  );
  const [showInHome, setShowInHome] = React.useState(false);
  const [newTitleError, setNewTitleError] = React.useState<string>("");
  const [newTitleWarning, setNewTitleWarning] = React.useState<string>(
    "Avoid special characters like #^*|\"':<>/?.",
  );
  const [selectedNewsItem, setSelectedNewsItem] = React.useState<any | null>(
    null,
  );
  const [isNewsDetailsOpen, setIsNewsDetailsOpen] = React.useState(false);

  const fetchNewsPost = async () => {
    const firstBatchSize = 5;
    setIsLoadingNews(true);
    try {
      const firstBatch = await props.provider.getNewsPost(
        props.searchSites,
        firstBatchSize,
        0,
      );
      console.log("First batch:", firstBatch);
      setNewsItems(firstBatch);

      setIsLoadingMoreNews(true);
      try {
        const remainingNews = await props.provider.getNewsPost(
          props.searchSites,
          200,
          firstBatchSize,
        );
        console.log("Remaining batch:", remainingNews);
        if (remainingNews.length > 0) {
          setNewsItems((prev) => [
            ...prev,
            ...remainingNews.filter(
              (item) => !prev.some((existing) => existing.Link === item.Link),
            ),
          ]);
        }
      } catch (error) {
        console.error("Error fetching remaining news posts:", error);
      } finally {
        setIsLoadingMoreNews(false);
      }
    } catch (error) {
      console.error("Error fetching news posts:", error);
      setNewsItems([]);
    } finally {
      setIsLoadingNews(false);
    }
  };
  const fetchHomeNewsPost = async () => {
    setIsLoadingNews(true);
    const noDepartmentFetch = await props.provider
      .getNewsPostWithNoDepartment(props.searchSites, 200, 0)
      .then((res) => {
        console.log(res);
        setNewsItems(res);
      })
      .catch((err) => {
        console.log("Error");
      })
      .finally(() => {
        setIsLoadingNews(false);
      });
  };
  React.useEffect(() => {
    if (props.layout == "Home") {
      fetchHomeNewsPost().catch((error) => {
        console.error("Error fetching home news posts:", error);
      });
    } else {
      fetchNewsPost().catch((error) => {
        console.error("Error fetching news posts:", error);
      });
    }
  }, []);
  const fetchDepartmentsOption = async () => {
    const news = await props.provider
      .getDepartmentFieldOptions(props.searchSites)
      .then((res) => {
        setDeptOptions(
          res.map((dept: string) => ({
            key: dept,
            text: dept,
            value: dept,
          })),
        );
      });
  };
  React.useEffect(() => {
    fetchDepartmentsOption().catch((error) => {
      console.error("Error fetching departments:", error);
    });
  }, []);
  // const fetchGroups = async () => {
  //   const graph = graphfi().using(graphSPFx(props.context));
  //   const groups = await graph.groups();
  //   const targetGroup: any = groups.find((g) => g.displayName === "DV Legal");

  //   if (targetGroup) {
  //     const gmembers = await graph.groups.getById(targetGroup.id).members();
  //     console.log(gmembers);
  //     const gowners = await graph.groups.getById(targetGroup.id).owners();
  //   }
  //   const members = await graph.groups
  //     .getById("4e52aa91-8ace-43c6-b006-e328fb1f0386")
  //     .owners();
  //   console.log(members);
  // };

  // React.useEffect(() => {
  //   fetchGroups().catch((error) => {
  //     console.error("Error fetching groups:", error);
  //   });
  // }, []);
  const responsive = React.useMemo(() => {
    const perPage = 4;
    return {
      // Extra small phones
      0: {
        slidesPerView: 1,
        spaceBetween: 10,
      },
      // Small phones / portrait
      480: {
        slidesPerView: Math.min(perPage, 2),
        spaceBetween: 15,
      },
      // Tablets
      768: {
        slidesPerView: Math.min(perPage, 3),
        spaceBetween: 20,
      },
      // Small desktops / laptops
      1024: {
        slidesPerView: Math.min(perPage, 4),
        spaceBetween: 20,
      },
      // Large desktops
      1440: {
        slidesPerView: perPage,
        spaceBetween: 20,
      },
    };
  }, [4]);

  const validateTitle = (value: string) => {
    const titleValue = value.trim();

    if (!titleValue) {
      setNewTitleError("Title is required.");
      return false;
    }

    const invalidTitlePattern = /[#^*|"'<>/:?\\]/;
    if (invalidTitlePattern.test(titleValue)) {
      setNewTitleError(
        "Title cannot contain special characters like #^*|\"':<>/?.",
      );
      return false;
    }

    setNewTitleError("");
    return true;
  };

  const handleCreateItem = async () => {
    const titleValue = newTitle.trim();
    if (!validateTitle(titleValue)) {
      return;
    }

    const createData = {
      title: titleValue,
      departments: selectedDepartments,
      showInHome,
    };

    setIsCreatingItem(true);
    // setIsAddDialogOpen(false);

    try {
      await props.provider
        .createNewsPost(createData, props.searchSites, props.sitetitle)
        .then((res: any) => {
          console.log(res);
          setRedirectUrl(res);
          setCreatedItemSuccess(true);
          setIsCreatingItem(false);
        });
    } catch (error) {
      console.error("Error creating news post:", error);
      setIsCreatingItem(false);
    }
  };

  const handleRedirect = () => {
    window.open(redirecturl, "_blank", "noopener,noreferrer");
    setCreatedItemSuccess(false);
    setIsAddDialogOpen(false);
  };

  const openAddDialog = () => {
    setNewTitle("");
    setSelectedDepartments([]);
    setSelectedSiteTypes([]);
    setShowInHome(false);
    setNewTitleError("");
    setCreatedItemSuccess(false);
    setIsAddDialogOpen(true);
  };

  const isAdminUser = props.adminUsers
    ? props.adminUsers.some(
        (user) => user.email?.toLowerCase() === props.userEmail?.toLowerCase(),
      )
    : [];

  const formatDetailDate = (value: string | undefined) => {
    if (!value) return "N/A";

    const parsedDate = moment(value);
    return parsedDate.isValid() ? parsedDate.format("DD MMMM YYYY") : "N/A";
  };

  const getDetailValue = (value: any) => {
    if (!value) return "N/A";

    if (Array.isArray(value)) {
      return value.filter(Boolean).join(", ");
    }

    if (typeof value === "object") {
      return value.Title || value.name || value.displayName || "N/A";
    }

    return String(value);
  };

  const getDepartmentValue = (item: any) => {
    const departmentValue =
      item?.Departments ||
      item?.Department ||
      item?.DepartmentName ||
      item?.DepartmentNames ||
      item?.DepartmentValues;

    return getDetailValue(departmentValue);
  };

  const openNewsDetails = (item: any) => {
    setSelectedNewsItem(item);
    setIsNewsDetailsOpen(true);
  };

  const closeNewsDetails = () => {
    setSelectedNewsItem(null);
    setIsNewsDetailsOpen(false);
  };

  const handleDepartmentChange = (_event: any, option?: any) => {
    if (!option) {
      return;
    }

    const optionKey = option.key as string;

    if (optionKey === "All") {
      setSelectedDepartments((prev) => (prev.includes("All") ? [] : ["All"]));
      return;
    }

    setSelectedDepartments((prev) => {
      if (prev.includes("All")) {
        return prev;
      }

      return prev.includes(optionKey)
        ? prev.filter((key) => key !== optionKey)
        : [...prev, optionKey];
    });
  };

  return (
    <>
      <div className={`${styles.featuredNews} `}>
        <div
          className={`${styles.contentBox}`}
          style={{
            borderTopColor: "black",
          }}
        >
          <div
            style={{ display: "flex", justifyContent: "space-between" }}
            className={styles.contentBox__heading}
          >
            <div className={`${styles.contentBox__heading}`}>
              <span className={styles.boxWebPartTitle}>
                {props.title || "News"}
              </span>
            </div>

            {isAdminUser && (
              <div className={styles.addForm}>
                <button
                  type="button"
                  className={styles.addFormBtn}
                  onClick={openAddDialog}
                >
                  Add
                </button>
              </div>
            )}
          </div>

          {isAddDialogOpen && (
            <div
              className={styles.modalOverlay}
              onClick={() => setIsAddDialogOpen(false)}
            >
              <div
                className={styles.modalContent}
                onClick={(event) => event.stopPropagation()}
              >
                <div className={styles.modalHeader}>
                  <div>
                    <div className={styles.modalTitle}>Add News Item</div>
                  </div>
                  <button
                    type="button"
                    className={styles.modalClose}
                    onClick={() => setIsAddDialogOpen(false)}
                    aria-label="Close"
                  >
                    ×
                  </button>
                </div>
                <div className={styles.modalBody}>
                  <div className={styles.formRow}>
                    <TextField
                      label="Title"
                      value={newTitle}
                      onChange={(event: any) => {
                        setNewTitle(event.target.value);
                        if (newTitleError) {
                          setNewTitleError("");
                        }
                      }}
                      placeholder="Enter title"
                      required
                      className="NewsTitle"
                      disabled={isCreatingItem}
                      description={newTitleWarning}
                      errorMessage={newTitleError}
                    />
                  </div>

                  <div className={styles.formRow}>
                    <div className={styles.inputLabelRow}>
                      <label className={styles.inputLabel}>Departments</label>
                      <button
                        type="button"
                        className={styles.clearBtn}
                        onClick={() => setSelectedDepartments([])}
                      >
                        Clear
                      </button>
                    </div>
                    <Dropdown
                      placeholder="Select Departments"
                      multiSelect
                      disabled={isCreatingItem}
                      className="deptdd"
                      selectedKeys={selectedDepartments}
                      onChange={handleDepartmentChange}
                      options={deptOptions.map((dept) => ({
                        ...dept,
                        disabled:
                          selectedDepartments.includes("All") &&
                          dept.key !== "All",
                      }))}
                    />
                  </div>

                  <div className={styles.toggleRow}>
                    {/* <Checkbox
                      label="Show in Home"
                      onChange={(event: any) =>
                        setShowInHome(event.target.checked)
                      }
                      disabled={isCreatingItem}
                    /> */}

                    {/* <label className={styles.checkboxField}>
                      <input
                        id="showInHome"
                        type="checkbox"
                        checked={showInHome}
                        onChange={(event) =>
                          setShowInHome(event.target.checked)
                        }
                      />
                      <span>Show in Home</span>
                    </label> */}
                  </div>
                </div>
                <div className={styles.buttonGroup}>
                  {/* <button
                    type="button"
                    className={styles.cancelBtn}
                    onClick={() => setIsAddDialogOpen(false)}
                  >
                    Cancel
                  </button> */}
                  <button
                    type="button"
                    className={styles.createBtn}
                    onClick={handleCreateItem}
                    disabled={isCreatingItem || !newTitle.trim()}
                  >
                    {isCreatingItem ? "Creating..." : "Create"}
                  </button>
                </div>
              </div>
            </div>
          )}
          {createdItemSuccess && (
            <>
              <div
                className={styles.modalOverlay}
                onClick={() => setCreatedItemSuccess(false)}
              >
                <div
                  className={styles.successModalContent}
                  onClick={(event) => event.stopPropagation()}
                >
                  <div className={styles.successIcon}>✓</div>
                  <div className={styles.successTitle}>
                    Item Created Successfully
                  </div>
                  <div className={styles.successMessage}>
                    Your news item has been created and will be available
                    shortly.
                  </div>
                  <button
                    type="button"
                    className={styles.redirectBtn}
                    onClick={handleRedirect}
                  >
                    Go to News Page
                  </button>
                </div>
              </div>
            </>
          )}
          {isNewsDetailsOpen && selectedNewsItem && (
            <div className={styles.modalOverlay} onClick={closeNewsDetails}>
              <div
                className={styles.newsDetailsModalContent}
                onClick={(event) => event.stopPropagation()}
              >
                <div className={styles.modalHeader}>
                  <div>
                    <div className={styles.modalTitle}>News Details</div>
                  </div>
                  <button
                    type="button"
                    className={styles.modalClose}
                    onClick={closeNewsDetails}
                    aria-label="Close"
                  >
                    ×
                  </button>
                </div>
                <div className={styles.modalBody}>
                  <div className={styles.detailRow}>
                    <span className={styles.detailLabel}>Title: </span>
                    <span className={styles.detailValue}>
                      {getDetailValue(selectedNewsItem.Title)}
                    </span>
                  </div>
                  <div className={styles.detailRow}>
                    <span className={styles.detailLabel}>Modified Date: </span>
                    <span className={styles.detailValue}>
                      {formatDetailDate(
                        selectedNewsItem.Modified ||
                          selectedNewsItem.modified ||
                          selectedNewsItem.LastModifiedTime ||
                          selectedNewsItem.date,
                      )}
                    </span>
                  </div>
                  <div className={styles.detailRow}>
                    <span className={styles.detailLabel}>
                      Modified Person:{" "}
                    </span>
                    <span className={styles.detailValue}>
                      {getDetailValue(
                        selectedNewsItem.AuthorByLine?.Title ||
                          selectedNewsItem.AuthorByLine?.EMail ||
                          selectedNewsItem.ModifiedBy ||
                          selectedNewsItem.Author ||
                          selectedNewsItem.AuthorByLine,
                      )}
                    </span>
                  </div>
                  {/* {props.layout !== "Home" && ( */}
                  <div className={styles.detailRow}>
                    <span className={styles.detailLabel}>Departments: </span>
                    <span className={styles.detailValue}>
                      {" "}
                      {getDepartmentValue(selectedNewsItem)}
                    </span>
                  </div>
                  {/* )} */}

                  <div className={styles.detailRow}>
                    <span className={styles.detailLabel}>Description: </span>
                    <div
                      className={styles.detailDescription}
                      dangerouslySetInnerHTML={{
                        __html:
                          selectedNewsItem.Description ||
                          "No description available.",
                      }}
                    />
                  </div>
                </div>
              </div>
            </div>
          )}
          <div>
            {isLoadingNews ? (
              <div className={styles.loadingContainer}>
                <Spinner label="Loading news..." size={SpinnerSize.large} />
              </div>
            ) : newsItems.length === 0 ? (
              <div className={styles.emptyState}>No results found.</div>
            ) : (
              <>
                <Swiper
                  // key={`${props.itemsToShowPerPage}-${props.item.length}`}
                  modules={[Navigation, Pagination]}
                  navigation={true}
                  // Base fallback; overridden by breakpoints for all widths >= 0
                  slidesPerView={1}
                  breakpoints={responsive}
                  className="featuredNews"
                >
                  {newsItems.length > 0 &&
                    newsItems.map((item: any, index: number) => {
                      const image = item.Link.includes("Pillars/SitePages")
                        ? `${
                            props.context.pageContext.site.absoluteUrl
                          }/_layouts/15/userphoto.aspx?size=L&username=${encodeURIComponent(
                            item.AuthorByLine && item.AuthorByLine.EMail
                              ? item.AuthorByLine.EMail
                              : "",
                          )}`
                        : item.Image;

                      return (
                        <SwiperSlide key={index}>
                          <div className={styles.filmstripView_Container}>
                            <div>
                              <img
                                src={image}
                                height={215}
                                alt={item.Title}
                                width="100%"
                              />
                              <div className={styles.titleDateContainer}>
                                <div className={styles.filmstripView_datetime}>
                                  {moment(item.date).format("DD MMMM YYYY")}
                                </div>
                                <button
                                  type="button"
                                  className={styles.filmstripView_titleButton}
                                  onClick={() => openNewsDetails(item)}
                                >
                                  {item.Title}
                                </button>
                                <div
                                  className={styles.DescriptionNews}
                                  title={item.Description}
                                  dangerouslySetInnerHTML={{
                                    __html: item.Description,
                                  }}
                                />
                                <a
                                  href={item.Link}
                                  target={"_blank"}
                                  data-interception={"propagate"}
                                  className={styles.learnMoreCont}
                                >
                                  {/* <span className={styles.learnMore}>
                                    {" "}
                                    Read More
                                  </span> */}
                                  <button
                                    onClick={() =>
                                      window.open(item.Link, "_blank")
                                    }
                                    className={styles.learnMore}
                                  >
                                    Read More
                                  </button>
                                </a>
                              </div>
                            </div>
                          </div>
                        </SwiperSlide>
                      );
                    })}
                </Swiper>
                {isLoadingMoreNews && (
                  <div className={styles.loadingContainer}>
                    <Spinner
                      label="Loading more news..."
                      size={SpinnerSize.small}
                    />
                  </div>
                )}
              </>
            )}
          </div>
        </div>
      </div>
    </>
  );
};
export default DvNews;
